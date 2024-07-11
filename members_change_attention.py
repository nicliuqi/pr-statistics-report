# 涉及到权限变更的PR，定期给相关审批人发送邮件提醒
from pr_statistics import *


def get_open_pulls():
    """
    Get a list of open pulls
    :return: a list of open pull requests
    """
    enterprise_pulls = []
    page = 1
    while True:
        log.logger.info("=" * 25 + " GET ENTERPRISE PULLS: PAGE {} ".format(page) + "=" * 25)
        url = 'https://gitee.com/api/v5/repos/openeuler/community/pulls'
        params = {
            'state': 'open',
            'sort': 'created',
            'direction': 'asc',
            'page': page,
            'per_page': 100,
            'access_token': os.getenv('ACCESS_TOKEN')
        }
        r = requests.get(url, params=params)
        if r.status_code != 200:
            log.logger.error('Fail to get enterprise pulls list.')
            return
        else:
            enterprise_pulls += r.json()
        if len(r.json()) < 100:
            break
        page += 1
    return enterprise_pulls


def pr_statistics(data_dir, open_pr_list):
    """
    :param data_dir: directory to store temporary data
    :param open_pr_list: open pull requests list
    """
    log.logger.info('=' * 25 + ' STATISTICS ' + '=' * 25)
    email_mappings = get_email_mappings()
    mapping_lists = sorted(list(email_mappings.keys()))
    open_pr_dict = {}
    open_pr_info = []
    for item in open_pr_list:
        if not item['mergeable'] or item['draft']:
            continue
        title = item['title']
        html_url = item['html_url']
        number = '#' + html_url.split('/')[-1]
        members = get_attention_members(html_url.split('/')[-1])
        if not members:
            continue
        created_at = item['created_at'].replace('T', ' ').replace('+08:00', '')
        labels = [x['name'] for x in item['labels']]
        ref_branch = item['head']['ref']
        status = '待合入'
        if 'openeuler-cla/yes' not in labels:
            status = fill_status(status, 'CLA认证失败')
        if 'ci_failed' in labels:
            status = fill_status(status, '门禁检查失败')
        if 'kind/wait_for_update' in labels:
            status = fill_status(status, '等待更新')
        duration = count_duration(created_at)
        link = "<a href='{0}'>{1}</a>".format(html_url, title)
        number_link = "<a href='{0}'>{1}</a>".format(html_url, number)
        open_pr_info.append(['TC', 'openeuler/community', ref_branch, number_link, link, status, duration,
                            ','.join(members)])

    no_addresses_id = []
    for pr_info in open_pr_info:
        ids = pr_info[-1]
        for i in ids.split(','):
            if i not in mapping_lists:
                if i not in no_addresses_id:
                    log.logger.warning('WARNING! gitee_id {} does not match any email address.'.format(i))
                    no_addresses_id.append(i)
            if i not in open_pr_dict.keys():
                open_pr_dict[i] = [pr_info[:-1]]
            else:
                open_pr_dict[i].append(pr_info[:-1])
    for receiver in sorted(list(open_pr_dict.keys())):
        origin_pr_list = sorted(open_pr_dict[receiver], key=(lambda x: int(x[6])), reverse=True)
        ordered_pr_list = []
        pr_sigs = sorted(set([x[0] for x in origin_pr_list]))
        for pr_sig in pr_sigs:
            for op in origin_pr_list:
                if op[0] == pr_sig:
                    if len(ordered_pr_list) > 0 and op[0] == ordered_pr_list[-1][0] and int(op[-1]) > \
                            int(ordered_pr_list[-1][-1]):
                        ordered_pr_list.insert(-1, op)
                    else:
                        ordered_pr_list.append(op)
        statistics_csv = '{}/statistics_{}.csv'.format(data_dir, receiver)
        f = codecs.open(statistics_csv, 'w', encoding='utf-8')
        writer = csv.writer(f)
        for i in ordered_pr_list:
            writer.writerow(i)
        f.close()
        email_address = email_mappings.get(receiver)
        if not email_address:
            log.logger.warning('Ready to send statistics for {} but cannot find the email address'.format(receiver))
            continue
        log.logger.info('Ready to send statistics for {} whose email address is {}'.format(receiver, email_address))
        statistics_xlsx = csv_to_xlsx(statistics_csv)
        excel_optimization(statistics_xlsx, {})
        send_email(statistics_xlsx, receiver, [email_address])


def excel_optimization(filepath, compare_dict):
    """
    Adjust styles of the xlsx file
    :param filepath: path of the xlsx file
    :param compare_dict: a dict of every sig and its compare info
    """
    if not filepath.endswith('.xlsx'):
        return
    html_file = filepath.replace('.xlsx', '.html')
    wb = openpyxl.load_workbook(filepath)
    ws = wb.active
    tmp_list = []
    for row in ws.rows:
        tmp_list.append(row[1].value)
    insert_rows = {}
    for i in tmp_list:
        if i not in insert_rows.keys():
            insert_row = tmp_list.index(i) + 1
            insert_rows[insert_row] = i
    # delete auxiliary column
    ws.delete_cols(1)
    ws.delete_cols(1)
    # insert rows
    alignment_center = Alignment(horizontal='center', vertical='center')
    insert_count = 0
    for i in sorted(insert_rows.keys()):
        sig_name = insert_rows[i]
        i += insert_count

        ws.insert_rows(i)
        insert_count += 1
        ws['A' + str(i)] = sig_name
        ws['A' + str(i)].font = Font(name='黑体', size=20, bold=True)
        ws['A' + str(i)].alignment = alignment_center

        compare_info = single_sig_compare(sig_name, compare_dict)
        ws.insert_rows(i + 1)
        insert_count += 1
        ws['A' + str(i + 1)] = compare_info
        ws['A' + str(i + 1)].font = Font(name='黑体', color='FF0000')
        ws['A' + str(i + 1)].alignment = alignment_center

        ws.insert_rows(i + 2)
        insert_count += 1
        ws['A' + str(i + 2)] = '仓库'
        ws['B' + str(i + 2)] = '目标分支'
        ws['C' + str(i + 2)] = '编号'
        ws['D' + str(i + 2)] = '标题'
        ws['E' + str(i + 2)] = '状态'
        ws['F' + str(i + 2)] = '开启天数'
        ws['A' + str(i + 2)].font = Font(bold=True)
        ws['A' + str(i + 2)].alignment = alignment_center
        ws['B' + str(i + 2)].font = Font(bold=True)
        ws['B' + str(i + 2)].alignment = alignment_center
        ws['C' + str(i + 2)].font = Font(bold=True)
        ws['C' + str(i + 2)].alignment = alignment_center
        ws['D' + str(i + 2)].font = Font(bold=True)
        ws['D' + str(i + 2)].alignment = alignment_center
        ws['E' + str(i + 2)].font = Font(bold=True)
        ws['E' + str(i + 2)].alignment = alignment_center
        ws['F' + str(i + 2)].font = Font(bold=True)
        ws['F' + str(i + 2)].alignment = alignment_center

    # replace the original table header
    ws.insert_rows(5)
    ws['A5'] = ws['A4'].value
    ws['B5'] = ws['B4'].value
    ws['C5'] = ws['C4'].value
    ws['D5'] = ws['D4'].value
    ws['E5'] = ws['E4'].value
    ws['F5'] = ws['F4'].value
    ws.delete_rows(4)
    # fill for the Duration
    cells = ws.iter_rows(min_row=3, min_col=6, max_col=6)
    yellow_fill = PatternFill("solid", start_color='FFFF00')
    first_stage_fill = PatternFill('solid', start_color='FFDAB9')
    second_stage_fill = PatternFill('solid', start_color='FF7F50')
    third_stage_fill = PatternFill('solid', start_color='FF4500')
    for i in cells:
        try:
            value = int(i[0].value)
            if 7 < value <= 30:
                i[0].fill = first_stage_fill
            elif 30 < value <= 365:
                i[0].fill = second_stage_fill
            elif value > 365:
                i[0].fill = third_stage_fill
        except (TypeError, ValueError):
            pass
    # fill for the status mark
    status = ws.iter_rows(min_row=3, min_col=5, max_col=5)
    for j in status:
        value = j[0].value
        if not value:
            continue
        elif len(value) <= 3 and value != '草稿':
            continue
        else:
            j[0].fill = yellow_fill
    # align center
    for row in ws.rows:
        row[5].alignment = alignment_center
    # add borders
    border = Border(left=Side(border_style='thin', color='000000'),
                    right=Side(border_style='thin', color='000000'),
                    top=Side(border_style='thin', color='000000'),
                    bottom=Side(border_style='thin', color='000000'))
    for row in ws.rows:
        for cell in row:
            cell.border = border
    ws.delete_rows(1)
    ws.delete_rows(1)
    ws.delete_cols(1)
    ws.delete_cols(1)
    wb.save(filepath)
    wb.close()
    # generate html file by the xlsx file
    xlsx2html(filepath, html_file)
    log.logger.info('Generate {}'.format(html_file))


def send_email(xlsx_file, nickname, receivers):
    """
    Send email to reviewers
    :param xlsx_file: path of the xlsx file
    :param nickname: Gitee ID of the receiver
    :param receivers: where send to
    """
    username = os.getenv('SMTP_USERNAME', '')
    port = os.getenv('SMTP_PORT', '')
    host = os.getenv('SMTP_HOST', '')
    password = os.getenv('SMTP_PASSWORD', '')
    sender = os.getenv('SMTP_SENDER')
    msg = MIMEMultipart()
    html_file = xlsx_file.replace('.xlsx', '.html')
    with open(html_file, 'r', encoding='utf-8') as f:
        body_of_email = f.read()
    body_of_email = body_of_email.replace(
        '<body>', '<body><p>Dear {},</p><p>以下是openEuler社区<b style="color:red">涉及SIG信息变更</b>的待处理PR，烦请您及时跟进</p>'.
            format(nickname)).replace('&nbsp;', '0')
    content = MIMEText(body_of_email, 'html', 'utf-8')
    msg.attach(content)
    msg['Subject'] = 'openEuler 成员变更待处理PR汇总'
    msg['From'] = sender
    msg['To'] = ','.join(receivers)
    try:
        if int(port) == 465:
            server = smtplib.SMTP_SSL(host, port)
            server.ehlo()
            server.login(username, password)
        else:
            server = smtplib.SMTP(host, port)
            server.ehlo()
            server.starttls()
            server.login(username, password)
        server.sendmail(sender, receivers, msg.as_string())
        log.logger.info('Sent report email to: {}'.format(receivers))
    except smtplib.SMTPException as e:
        log.logger.error(e)


def get_all_comments(number):
    all_comments = []
    page = 1
    while True:
        url = 'https://gitee.com/api/v5/repos/openeuler/community/pulls/{}/comments'.format(number)
        params = {
            'page': page,
            'per_page': 100,
            'access_token': os.getenv('ACCESS_TOKEN')
        }
        r = requests.get(url, params=params)
        if r.status_code != 200:
            break
        if len(r.json()) == 0:
            break
        all_comments.extend(r.json())
        page += 1
    return all_comments


def get_pr_lgtm_list(all_comments):
    pr_lgtm_list = []
    review_key = '以下为 openEuler-Advisor 的 review_tool 生成审视要求清单'
    review_key_en = 'The following table is the PR review checklist generated by the review_tool of openEuler-Advisor'
    latest_review_comment = None
    for comment in reversed(all_comments):
        if comment['user']['login'] != 'openeuler-ci-bot':
            continue
        if review_key in comment['body'] or review_key_en in comment['body']:
            latest_review_comment = comment
    if not latest_review_comment:
        return pr_lgtm_list

    review_checklist = latest_review_comment['body']

    for i in review_checklist.split('\n'):
        if len(i.split('|')) != 7:
            continue
        if i.split('|')[5] not in ['[&#x1F534;]', '[&#x25EF;]', '[&#x1F7E1;]', '[&#x1F535;]']:
            continue
        for j in i.split('|')[4].split(' '):
            if j.startswith('@') and j[1:] not in pr_lgtm_list:
                pr_lgtm_list.append(j[1:])
    return pr_lgtm_list


def get_attention_members(number):
    all_comments = get_all_comments(number)
    return get_pr_lgtm_list(all_comments)


if __name__ == '__main__':
    data_dir = prepare_env()
    open_pr_list = get_open_pulls()
    pr_statistics(data_dir, open_pr_list)


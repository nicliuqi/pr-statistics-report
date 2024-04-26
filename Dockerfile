FROM openeuler/openeuler:22.03-lts

MAINTAINER liuqi<469227928@qq.com>

RUN yum install -y python3-pip git

RUN pip3 install requests openpyxl pandas PyYAML xlsx2html -i https://pypi.tuna.tsinghua.edu.cn/simple

WORKDIR /work/pr-statistics

COPY . /work/pr-statistics

ENV TZ=Asia/Shanghai

ENTRYPOINT ["python3", "pr_statistics.py"] 

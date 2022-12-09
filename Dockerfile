FROM openeuler/openeuler:21.03

MAINTAINER liuqi<469227928@qq.com>

RUN yum update && \
yum install -y python3-pip git

RUN pip3 install requests openpyxl pandas PyYAML xlsx2html

WORKDIR /work/pr-statistics

COPY pr_statistics.py /work/pr-statistics

ENTRYPOINT ["python3", "pr_statistics.py"] 

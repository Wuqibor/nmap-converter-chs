#!/usr/bin/env python3

from libnmap.parser import NmapParser, NmapParserException
from xlsxwriter import Workbook
from datetime import datetime

import os.path


class HostModule:
    def __init__(self, host):
        self.host = next(iter(host.hostnames), "")
        self.ip = host.address
        self.port = ""
        self.protocol = ""
        self.status = ""
        self.service = ""
        self.tunnel = ""
        self.method = ""
        self.source = ""
        self.confidence = ""
        self.reason = ""
        self.reason = ""
        self.product = ""
        self.version = ""
        self.extra = ""
        self.flagged = "N/A"
        self.notes = ""


class ServiceModule(HostModule):
    def __init__(self, host, service):
        super(ServiceModule, self).__init__(host)
        self.host = next(iter(host.hostnames), "")
        self.ip = host.address
        self.port = service.port
        self.protocol = service.protocol
        self.status = service.state
        self.service = service.service
        self.tunnel = service.tunnel
        self.method = service.service_dict.get("method", "")
        self.source = "scanner"
        self.confidence = float(service.service_dict.get("conf", "0")) / 10
        self.reason = service.reason
        self.product = service.service_dict.get("product", "")
        self.version = service.service_dict.get("version", "")
        self.extra = service.service_dict.get("extrainfo", "")
        self.flagged = "N/A"
        self.notes = ""


class HostScriptModule(HostModule):
    def __init__(self, host, script):
        super(HostScriptModule, self).__init__(host)
        self.method = script["id"]
        self.source = "script"
        self.extra = script["output"].strip()


class ServiceScriptModule(ServiceModule):
    def __init__(self, host, service, script):
        super(ServiceScriptModule, self).__init__(host, service)
        self.source = "script"
        self.method = script["id"]
        self.extra = script["output"].strip()


def _tgetattr(object, name, default=None):
    try:
        return getattr(object, name, default)
    except Exception:
        return default


def generate_summary(workbook, sheet, report):
    summary_header = ["扫描", "命令", "版本", "扫描类型", "开始时间", "完成时间", "目标总数", "开启(ping)", "关闭(ping)"]
    sheet.freeze_panes(1, 0)
    summary_body = {"扫描": lambda report: _tgetattr(report, 'basename', 'N/A'),
                    "命令": lambda report: _tgetattr(report, 'commandline', 'N/A'),
                    "版本": lambda report: _tgetattr(report, 'version', 'N/A'),
                    "扫描类型": lambda report: _tgetattr(report, 'scan_type', 'N/A'),
                    "开始时间": lambda report: datetime.utcfromtimestamp(_tgetattr(report, 'started', 0)).strftime("%Y-%m-%d %H:%M:%S (UTC)"),
                    "完成时间": lambda report: datetime.utcfromtimestamp(_tgetattr(report, 'endtime', 0)).strftime("%Y-%m-%d %H:%M:%S (UTC)"),
                    "目标总数": lambda report: _tgetattr(report, 'hosts_total', 'N/A'),
                    "开启(ping)": lambda report: _tgetattr(report, 'hosts_up', 'N/A'),
                    "关闭(ping)": lambda report: _tgetattr(report, 'hosts_down', 'N/A')}

    for idx, item in enumerate(summary_header):
        sheet.write(0, idx, item, workbook.myformats["fmt_bold"])
        for idx, item in enumerate(summary_header):
            sheet.write(sheet.lastrow + 1, idx, summary_body[item](report))

    sheet.lastrow = sheet.lastrow + 1


def generate_hosts(workbook, sheet, report):
    sheet.autofilter("A1:E1")
    sheet.freeze_panes(1, 0)
    hosts_header = ["域名", "IP", "状态", "服务", "系统"]
    hosts_body = {"域名": lambda host: next(iter(host.hostnames), ""),
                  "IP": lambda host: host.address,
                  "状态": lambda host: host.status,
                  "服务": lambda host: len(host.services),
                  "系统": lambda host: os_class_string(host.os_class_probabilities())}

    for idx, item in enumerate(hosts_header):
        sheet.write(0, idx, item, workbook.myformats["fmt_bold"])

    row = sheet.lastrow
    for host in report.hosts:
        for idx, item in enumerate(hosts_header):
            sheet.write(row + 1, idx, hosts_body[item](host))
        row += 1

    sheet.lastrow = row


def generate_results(workbook, sheet, report):
    sheet.autofilter("A1:N1")
    sheet.freeze_panes(1, 0)
    results_header = ["域名", "IP", "端口", "协议", "状态", "服务", "通道", "来源", "方法", "准确性", "原因", "应用", "版本", "附加", "标记", "记录"]
    results_body = {"域名": lambda module: module.host,
                    "IP": lambda module: module.ip,
                    "端口": lambda module: module.port,
                    "协议": lambda module: module.protocol,
                    "状态": lambda module: module.status,
                    "服务": lambda module: module.service,
                    "通道": lambda module: module.tunnel,
                    "来源": lambda module: module.source,
                    "方法": lambda module: module.method,
                    "准确性": lambda module: module.confidence,
                    "原因": lambda module: module.reason,
                    "应用": lambda module: module.product,
                    "版本": lambda module: module.version,
                    "附加": lambda module: module.extra,
                    "标记": lambda module: module.flagged,
                    "记录": lambda module: module.notes}

    results_format = {"Confidence": workbook.myformats["fmt_conf"]}

    print("[+] 处理 {}".format(report.summary))
    for idx, item in enumerate(results_header):
        sheet.write(0, idx, item, workbook.myformats["fmt_bold"])

    row = sheet.lastrow
    for host in report.hosts:
        print("[+] 处理 {}".format(host))

        for script in host.scripts_results:
            module = HostScriptModule(host, script)
            for idx, item in enumerate(results_header):
                sheet.write(row + 1, idx, results_body[item](module), results_format.get(item, None))
            row += 1

        for service in host.services:
            module = ServiceModule(host, service)
            for idx, item in enumerate(results_header):
                sheet.write(row + 1, idx, results_body[item](module), results_format.get(item, None))
            row += 1

            for script in service.scripts_results:
                module = ServiceScriptModule(host, service, script)
                for idx, item in enumerate(results_header):
                    sheet.write(row + 1, idx, results_body[item](module), results_format.get(item, None))
                row += 1

    sheet.data_validation("O2:O${}".format(row + 1), {"validate": "list",
                                                      "source": ["Y", "N", "N/A"]})
    sheet.lastrow = row


def setup_workbook_formats(workbook):
    formats = {"fmt_bold": workbook.add_format({"bold": True}),
               "fmt_conf": workbook.add_format()}

    formats["fmt_conf"].set_num_format("0%")
    return formats


def os_class_string(os_class_array):
    return " | ".join(["{0} ({1}%)".format(os_string(osc), osc.accuracy) for osc in os_class_array])


def os_string(os_class):
    rval = "{0}, {1}".format(os_class.vendor, os_class.osfamily)
    if len(os_class.osgen):
        rval += "({0})".format(os_class.osgen)
    return rval


def main(reports, workbook):
    sheets = {"总览": generate_summary,
              "目标": generate_hosts,
              "结果": generate_results}

    workbook.myformats = setup_workbook_formats(workbook)

    for sheet_name, sheet_func in sheets.items():
        sheet = workbook.add_worksheet(sheet_name)
        sheet.lastrow = 0
        for report in reports:
            sheet_func(workbook, sheet, report)
    workbook.close()


if __name__ == '__main__':

    import argparse

    parser = argparse.ArgumentParser()
    parser.add_argument('-r', '--reports', required=True, nargs="+", help='nmap 扫描后输出的 xml 文件的路径')
    parser.add_argument('-o', '--output', help='输出转换后的 xlsx 文件的路径')
    args = parser.parse_args()

    xml_reports = []
    for tmp_path in args.reports:
        if tmp_path.endswith('.xml') and os.path.isfile(tmp_path):
            xml_reports.append(tmp_path)
        elif os.path.isdir(tmp_path):
            for file_path in os.listdir(tmp_path):
                if file_path.endswith('.xml'):
                    xml_reports.append(os.path.join(tmp_path, file_path))
        else:
            parser.print_help()
            print(f'\n[!] "{tmp_path}" 不是文件或目录')
            exit()

    reports = []
    for report in xml_reports:
        try:
            parsed = NmapParser.parse_fromfile(report)
        except NmapParserException as e:
            parsed = NmapParser.parse_fromfile(report, incomplete=True)

        setattr(parsed, 'source', os.path.basename(report))
        reports.append(parsed)

    xlsx_path = args.output if args.output else f'Report_%s' % datetime.now().strftime('%Y%m%d_%H%M%S')
    if not xlsx_path.endswith('.xlsx'):
        xlsx_path += '.xlsx'
    workbook = Workbook(xlsx_path)
    main(reports, workbook)

    print("感谢使用 Nmap-Converter-CHS")
    print("项目源作者: https://github.com/mrschyte/nmap-converter")
    print("参考项目: https://github.com/0xn0ne/nmapReport")
    print("中文修改(此)作者: https://github.com/Wuqibor/nmap-converter-chs")
    print("转换后的文件已经保存至 %s" % xlsx_path)

## 报表插件
报表插件提供快速生成在线Html5报表页面，并提供报表下载及分页静默打印功能。
* 报表插件1.0.0-SNAPSHOT目前支持Spring Boot 2.0.2.RELEASE版本
* 报表插件1.0.1-SNAPSHOT目前支持Spring Boot 2.1.2.RELEASE版本

###  如何使用
1.在pom文件中引入依赖，目前最新版本为：1.0.0-SNAPSHOT
```
    <dependency>
        <groupId>com.wk.core.plugins</groupId>
        <artifactId>wk-core-plugins-report</artifactId>
        <version>0.0.1</version>
    </dependency>
```
2.在项目properties/ymal配置文件中添加如下配置:

````
.perperties配置如下：
# freemarker配置
spring.freemarker.request-context-attribute=req
spring.freemarker.suffix=.html
spring.freemarker.content-type=text/html
spring.freemarker.enabled=true
spring.freemarker.cache=false
spring.freemarker.template-loader-path=classpath:/templates/
spring.freemarker.charset=UTF-8
spring.freemarker.settings.number_format=0.##

#项目模板配置
# tpl模板存放路径，tpl模板可在集团项目脚手架获取
zdxf.report.tpl=D:/tmp/report/tpl
# excel模板存放路径
zdxf.report.template=D:/tmp/report/template
# 导出Excel临时目录
zdxf.report.genPath=D:/tmp/report/download

````

````
.yml配置如下：
spring:
  freemarker:
    cache: false
    charset: UTF-8
    content-type: text/html
    enabled: true
    request-context-attribute: req
    settings:
      number_format: 0.##
    suffix: .html
    template-loader-path: classpath:/templates/
zdxf:
  report:
    genPath: D:/tmp/report/download
    template: D:/tmp/report/template
    tpl: D:/tmp/report/tpl
````
3.后台调用create接口生成报表数据
````
http://localhost:8850/XXX/report/create
该接口为POST接口，请求参数如下
{    
    "fileName":"ticket",
    "headerValues":[
        {
            "label":"统计时间:",
            "value":"2020-01-01 00:00:00--2020-11-30 16:45:40"
        },
        {
            "label":"售票员:",
            "value":"标准测试"
        },
        {
            "label":"打印时间:",
            "value":"2020-11-30 16:45:40"
        }
    ],
    "pageSize":13,
    "headerRows":4,
    "showIndexs":[
        0,
        1,
        2,
        3,
        4,
        5,
        6,
        7,
        8,
        9
    ],
"values":[
        "散客票/rp=4;奇梁洞门票;标准票;5.00;5;0;5;25.00;0.00;25.00",
        " 散客票;app车票0818;儿童票;20.00;1;0;1;20.00;0.00;20.00",
        " 散客票;接驳车通票;标准票;10.00;4;0;4;40.00;0.00;40.00",
        " 散客票;测试0826;标准票;30.00;2;0;2;130.00;0.00;130.00",
        " 团体票;小计/cp=3;_;_;6;0;6;56.00;0.00;56.00",
        " 合计1/rp=5;奇梁洞门票;标准票;5.00;5;0;5;25.00;0.00;25.00",
        " 合计;app车票0818;儿童票;20.00;1;0;1;20.00;0.00;20.00",
        " 合计;接驳车通票;标准票;10.00;4;0;4;40.00;0.00;40.00",
        " 合计;测试0826;标准票;30.00;2;0;2;130.00;0.00;130.00",
        " 合计;舞蹈芭蕾;标准票;2.00;1;1;0;2.00;2.00;0.00",
        " 现金/rp=3;现金收款(￥)/cp=3;_;_;716.06/cp=6;",
        " 现金;现金退款(￥)/cp=3;_;_;142.00/cp=6;",
        " 现金;现金交付(￥)/cp=3;_;_;574.06/cp=6;",
        " 身份证;小计/cp=3;_;_;18;1;17;155.11;100.00;55.11",
        " 纸质票;小计/cp=3;_;_;55;3;52;571.05;42.00;529.05"
    ]
}
````
参数数据说明
````
fileName：       （必填）excel模板文件名也是导出后的文件名，导出后文件名为fileName+时间戳.xls
headerValues：   （可填）报表头部自定义信息
pageSize:       (必填）报表每页行数，包含模板的表头和数据行数，不包含自定义表头。
headerRows：     （必填）模板表头终止行的行标，即NA上级行号。
showIndexs：     （可填）显示列信息，若不填则显示模板配置的所有列数据。
values:         （必填）报表业务数据。
````
values配置规则
````
values为字符数组，每行用逗号分割。每行内每个单元格数据用分号分割。单元格跨行用/rp标识，比如散客票/rp=4即散客票夸4行。跨列用cp标识，比如小计/cp=3即小计夸3列。
若表头需要斜线用\\配置，如excel模板中配置项目\\名称即生成斜线分割。在values数据层若导出需要斜线则须转义，用\\\\配置。

````


    

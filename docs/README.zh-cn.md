# [simter-jxls-ext](https://github.com/simter/simter-jxls-ext) [[English]]

Simter [Jxls] Extensions. Includes:
- Common Functions
    - Format number: fn.format(1123456789.456, '#,###.00') > 1,123,456,789.56
    - Round number: fn.round(45.678, 2) > 45.68
    - Convert string to int: fn.toInt('123') > 123
    - Concat multipul strings: fn.concat('ab', 'c', ...) > 'abc'
    - Format java8 date/time: fn.format(LocalDateTime.now(), 'yyyy-MM-dd HH:mm:ss') > 2017-01-01 12:30:50

- EachMergeCommand: `jx:each-merge`, for auto merge cells

See the usage code bellow.

## 安装

```xml
<dependency>
  <groupId>tech.simter</groupId>
  <artifactId>simter-jxls-ext</artifactId>
  <version>0.3.0</version>
</dependency>
```

## 要求

- [Jxls] 2+
- Java 8+

## 使用

### Common Functions

```java
Context context = new Context();

// inject common-functions
context.putVar("fn", CommonFunctions.getSingleton());

// other data
context.putVar("num", new BigDecimal("123.456"));
context.putVar("datetime", LocalDateTime.now());
context.putVar("date", LocalDate.now());
context.putVar("time", LocalTime.now());
context.putVar("str", "123");

// render template
InputStream template = ...;
OutputStream output = ...;
JxlsHelper.getInstance().processTemplate(template, output, context);
```

相应的单元测试代码参见 [CommonFunctionsTest.java]. [Excel 模板][common-functions-template] 和渲染结果截图如下:
![common-functions.png]

### EachMergeCommand - 自动合并单元格

```java
// 全局注册自定义的 each-merge 指令到 XlsCommentAreaBuilder
XlsCommentAreaBuilder.addCommandMapping(EachMergeCommand.COMMAND_NAME, EachMergeCommand.class);

// 生成 "主-从" 结构的数据
Context context = new Context();
context.putVar("rows", generateRowsData());

// 渲染模板
InputStream template = ...;
OutputStream output = ...;
JxlsHelper.getInstance().processTemplate(template, output, context);
```

方法 `generateRowsData()` 生成的数据结构如下:

```javascript
[
  {
    sn: 1, 
    name: 'row1',
    subs: [
      {sn: '1-1', name: 'row1sub1'},
      ...
    ]
  },
  ...
]
```

相应的单元测试代码参见 [EachMergeCommandTest.java]. [Excel 模板][each-merge-template] 和渲染结果截图如下:
![each-merge.png]

## 构建

```bash
mvn clean package
```

## 发布

请先查看 [simter-parent] 的发布配置说明。

### 发布到局域网 Nexus 仓库

```bash
mvn clean deploy -P lan
```

### 发布到 Sonatype 仓库

```bash
mvn clean deploy -P sonatype
```

发布成功后登陆到 <https://oss.sonatype.org>，在 `Staging Repositories` 找到这个包，然后将其 close 和 release。
过几个小时后，就会自动同步到 [Maven 中心仓库](http://repo1.maven.org/maven2/tech/simter/simter-jxls-ext) 了。

### 发布到 Bintray 仓库

```bash
mvn clean deploy -P bintray
```

发布之前要先在 Bintray 创建 package `https://bintray.com/simter/maven/tech.simter:simter-jxls-ext`。
发布到的地址为 `https://api.bintray.com/maven/simter/maven/tech.simter:simter-jxls-ext/;publish=1`。
发布成功后可以到 <https://jcenter.bintray.com/tech/simter/simter-jxls-ext> 检查一下结果。


[English]: https://github.com/simter/simter-jxls-ext/blob/master/README.md
[simter-parent]: https://github.com/simter/simter-parent/blob/master/docs/README.zh-cn.md

[Jxls]: http://jxls.sourceforge.net
[oss.sonatype.org]: https://oss.sonatype.org

[CommonFunctionsTest.java]: https://github.com/simter/simter-jxls-ext/blob/master/src/test/java/tech/simter/jxls/ext/CommonFunctionsTest.java#L77
[common-functions-template]: https://github.com/simter/simter-jxls-ext/raw/master/src/test/resources/templates/common-functions.xlsx
[common-functions.png]: common-functions.png

[EachMergeCommandTest.java]: https://github.com/simter/simter-jxls-ext/blob/master/src/test/java/tech/simter/jxls/ext/EachMergeCommandTest.java#L30
[each-merge-template]: https://github.com/simter/simter-jxls-ext/raw/master/src/test/resources/templates/each-merge.xlsx
[each-merge.png]: each-merge.png
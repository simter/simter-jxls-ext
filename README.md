# [simter-jxls-ext](https://github.com/simter/simter-jxls-ext) [[中文]]

Simter [Jxls] Extensions. Includes:
- Common Functions
    - Format number: fn.format(1123456789.456, '#,###.00') > 1,123,456,789.56
    - Round number: fn.round(45.678, 2) > 45.68
    - Convert string to int: fn.toInt('123') > 123
    - Concat multipul strings: fn.concat('ab', 'c', ...) > 'abc'
    - Format java8 date/time: fn.format(LocalDateTime.now(), 'yyyy-MM-dd HH:mm:ss') > 2017-01-01 12:30:50

- EachMergeCommand: `jx:each-merge`, for auto merge cells

See the usage code bellow.

## Installation

```xml
<dependency>
  <groupId>tech.simter</groupId>
  <artifactId>simter-jxls-ext</artifactId>
  <version>1.0.0</version>
</dependency>
```

## Requirement

- [Jxls] 2+
- Java 8+

## Usage

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

Check the unit test code from [CommonFunctionsTest.java]. The [template][common-functions-template] and render result show bellow:

![common-functions.png]

### EachMergeCommand - Auto merge cells

```java
// global add custom each-merge command to XlsCommentAreaBuilder
XlsCommentAreaBuilder.addCommandMapping(EachMergeCommand.COMMAND_NAME, EachMergeCommand.class);

// generate a main-sub structure data
Context context = new Context();
context.putVar("rows", generateRowsData());

// render template
InputStream template = ...;
OutputStream output = ...;
JxlsHelper.getInstance().processTemplate(template, output, context);
```

The `generateRowsData()` method generates the bellow structure data:

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

Check the unit test code from [EachMergeCommandTest.java]. The [template][each-merge-template] and render result show bellow:

![each-merge.png]

## Build

```bash
mvn clean package
```

## Deploy

First take a look at [simter-parent] deploy config.

### Deploy to LAN Nexus Repository

```bash
mvn clean deploy -P lan
```

### Deploy to Sonatype Repository

```bash
mvn clean deploy -P sonatype
```

After deployed, login into <https://oss.sonatype.org>. Through `Staging Repositories`, search this package, 
then close and release it. After couple hours, it will be synced 
to [Maven Central Repository](http://repo1.maven.org/maven2/tech/simter/simter-jxls-ext).

### Deploy to Bintray Repository

```bash
mvn clean deploy -P bintray
```

Will deploy to `https://api.bintray.com/maven/simter/maven/tech.simter:simter-jxls-ext/;publish=1`.
So first create a package `https://bintray.com/simter/maven/tech.simter:simter-jxls-ext` on Bintray.
After deployed, check it from <https://jcenter.bintray.com/tech/simter/simter-jxls-ext>.


[Jxls]: http://jxls.sourceforge.net
[oss.sonatype.org]: https://oss.sonatype.org
[simter-parent]: https://github.com/simter/simter-parent
[中文]: https://github.com/simter/simter-jxls-ext/blob/master/docs/README.zh-cn.md

[CommonFunctionsTest.java]: https://github.com/simter/simter-jxls-ext/blob/master/src/test/java/tech/simter/jxls/ext/CommonFunctionsTest.java#L77
[common-functions-template]: https://github.com/simter/simter-jxls-ext/raw/master/src/test/resources/templates/common-functions.xlsx
[common-functions.png]: docs/common-functions.png

[EachMergeCommandTest.java]: https://github.com/simter/simter-jxls-ext/blob/master/src/test/java/tech/simter/jxls/ext/EachMergeCommandTest.java#L30
[each-merge-template]: https://github.com/simter/simter-jxls-ext/raw/master/src/test/resources/templates/each-merge.xlsx
[each-merge.png]: docs/each-merge.png
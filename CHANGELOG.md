# simter-jxls-ext changelog

## 2.0.0-M1 - 2020-06-02

- Upgrade to simter-dependencies-2.0.0-M1

## 1.2.0-M3 - 2020-04-15

- Upgrade to simter-1.3.0-M14

## 1.2.0-M2 - 2020-02-10

- Upgrade to simter-1.3.0-M13

## 1.2.0-M1 - 2020-01-08

- Upgrade to simter-1.3.0-M11
- Support join list item to string with a special delimiter
    - Add `String CommonFunctions.join(List<Object> list, String delimiter)` method
    - Add `String CommonFunctions.join(List<Object> list)` method
- Support join special key or property value of list item to a string with a special delimiter
    - Add `String CommonFunctions.joinProperty(List<Object> list, String name, String delimiter)` method
    - Add `String CommonFunctions.joinProperty(List<Object> list, String namer)` method
- Support format `Duration`
    - Add `String duration(Temporal startTime, Temporal endTime)` method
    - Add `String format(Duration duration)` method

## 1.1.0 - 2019-07-03

No code changed, just polishing maven config and unit test.

- Use JUnit5|AssertJ instead of JUnit4|Hamcrest
- Change parent to simter-dependencies-1.2.0

## 1.0.0 - 2019-01-08

- Just align version

## 0.5.0 - 2018-01-05

- Just align version

## 0.4.0 - 2018-01-05

- Just centralize-version

## 0.3.0 - 2017-12-12

- Add Jxls common functions
- Add Jxls `jx:each-merge` command for auto merge cells
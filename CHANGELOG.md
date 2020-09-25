# simter-jxls-ext changelog

## 1.1.1 - 2020-09-25

- Use afterApplyAtCell instead of afterTransformCell in EachMergeCommand

> For jxls-2.6 parent's afterTransformCell is call after sub's afterTransformCell.
> But change after jxls-2.7+. 
> This change is for future compatibility.

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
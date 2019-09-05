package tech.simter.jxls.ext;

import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.time.Duration;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.time.temporal.TemporalAccessor;
import java.util.Arrays;
import java.util.stream.Collectors;

/**
 * The common functions for jxls.
 * <p>
 * 1) format date-time.<br>
 * 2) format number.<br>
 * 3) concat string.<br>
 * 4) string to int.<br>
 *
 * @author RJ
 */
public final class CommonFunctions {
  private static CommonFunctions singleton = new CommonFunctions();

  public static CommonFunctions getSingleton() {
    return singleton;
  }

  private CommonFunctions() {
  }

  /**
   * java.time 的格式化。
   *
   * @param temporal 值
   * @param pattern  格式
   * @return 格式化后的值
   */
  public String format(TemporalAccessor temporal, String pattern) {
    return temporal == null ? null : DateTimeFormatter.ofPattern(pattern).format(temporal);
  }

  /**
   * 数字的格式化。
   *
   * @param number  值
   * @param pattern 格式
   * @return 格式化后的值
   */
  public String format(Number number, String pattern) {
    return number == null ? null : new DecimalFormat(pattern).format(number);
  }

  /**
   * 数字的四舍五入。
   *
   * @param number 值
   * @param scale  小数位数
   * @return 四舍五入后的值
   */
  public Number round(Number number, int scale) {
    return number == null ? null : (number instanceof BigDecimal ?
      ((BigDecimal) number).setScale(scale, BigDecimal.ROUND_HALF_UP)
      : new BigDecimal(number.toString()).setScale(scale, BigDecimal.ROUND_HALF_UP));
  }

  /**
   * 合并字符串
   *
   * @param items 合并项
   * @return 合并后的字符串
   */
  public String concat(Object... items) {
    if (items == null || items.length == 0) return null;
    return Arrays.stream(items).map(i -> i == null ? "" : i.toString()).collect(Collectors.joining());
  }

  /**
   * 转换为整数
   *
   * @param str 字符值
   * @return 整数
   */
  public Integer toInt(String str) {
    return str == null ? null : new Integer(str);
  }

  /**
   * 计算两个时间差，endTime 必须比 startTime 大
   *
   * @param startTime 开始时间
   * @param endTime   结束时间
   * @return 格式化的字符串，格式如：10d10h10m
   */
  public String timeSubtract(LocalDateTime startTime, LocalDateTime endTime) {
    if (startTime == null || endTime == null) {
      return null;
    }
    Duration duration = Duration.between(startTime, endTime);
    StringBuffer formatTime = new StringBuffer();
    if (duration.toDays() > 0) {
      formatTime.append(duration.toDays() + "d");
    }
    if (duration.toHours() % 24 > 0) {
      formatTime.append((duration.toHours() % 24) + "h");
    }
    if (duration.toMinutes() % 60 > 0) {
      formatTime.append((duration.toMinutes() % 60) + "m");
    }
    return formatTime.toString();
  }
}

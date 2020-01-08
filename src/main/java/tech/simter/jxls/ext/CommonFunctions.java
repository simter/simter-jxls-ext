package tech.simter.jxls.ext;

import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.time.format.DateTimeFormatter;
import java.time.temporal.TemporalAccessor;
import java.util.Arrays;
import java.util.List;
import java.util.Objects;
import java.util.stream.Collectors;

/**
 * The common functions for jxls.
 * <p>
 * 1) format date-time.<br>
 * 2) format number.<br>
 * 3) concat string.<br>
 * 4) string to int.<br>
 * 5) join list to string.<br>
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
   * Format java.time.
   *
   * @param temporal javaTime instance
   * @param pattern  the formatter
   * @return the formatted value
   */
  public String format(TemporalAccessor temporal, String pattern) {
    return temporal == null ? null : DateTimeFormatter.ofPattern(pattern).format(temporal);
  }

  /**
   * Format number.
   *
   * @param number  the number value
   * @param pattern the formatter
   * @return the formatted value
   */
  public String format(Number number, String pattern) {
    return number == null ? null : new DecimalFormat(pattern).format(number);
  }

  /**
   * Round number half up with special scale.
   *
   * @param number the number value
   * @param scale  the scale value
   * @return half up number value
   */
  public Number round(Number number, int scale) {
    return number == null ? null : (number instanceof BigDecimal ?
      ((BigDecimal) number).setScale(scale, BigDecimal.ROUND_HALF_UP)
      : new BigDecimal(number.toString()).setScale(scale, BigDecimal.ROUND_HALF_UP));
  }

  /**
   * Concat all param to a string.
   *
   * @param items the params
   * @return a string
   */
  public String concat(Object... items) {
    if (items == null || items.length == 0) return null;
    return Arrays.stream(items).map(i -> i == null ? "" : i.toString()).collect(Collectors.joining());
  }

  /**
   * Convert string value to a Integer value.
   *
   * @param str the string value
   * @return an Integer value
   */
  public Integer toInt(String str) {
    return str == null ? null : new Integer(str);
  }

  /**
   * Join all list item to a string with a special delimiter.
   * <p>
   * Note：null item would be ignored.
   *
   * @param list      the list to join
   * @param delimiter the delimiter to join item
   * @return a joined string
   */
  public String join(List<Object> list, String delimiter) {
    if (delimiter == null) delimiter = ", ";
    return list.stream()
      .filter(Objects::nonNull)
      .map(Object::toString)
      .collect(Collectors.joining(delimiter));
  }

  /**
   * Join all list item to a string with ", " delimiter.
   * <p>
   * Note：null item would be ignored.
   *
   * @param list the list to join
   * @return a joined string
   */
  public String join(List<Object> list) {
    return join(list, ", ");
  }
}

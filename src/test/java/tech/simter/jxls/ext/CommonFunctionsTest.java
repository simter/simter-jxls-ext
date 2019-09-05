package tech.simter.jxls.ext;

import org.junit.jupiter.api.Test;
import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.time.*;
import java.time.temporal.TemporalAccessor;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;

/**
 * Common Functions test.
 *
 * @author RJ
 */
class CommonFunctionsTest {
  private final CommonFunctions fn = CommonFunctions.getSingleton();

  @Test
  void formatNumber() {
    assertNull(fn.round(null, 0));
    assertEquals("123,456,789.10", fn.format(123456789.1, "#,###.00"));
    assertEquals("123,456,789.12", fn.format(123456789.12, "#,###.00"));
    assertEquals("123,456,789.12", fn.format(123456789.124, "#,###.00"));
    assertEquals("123,456,789.12", fn.format(123456789.125, "#,###.00"));
    assertEquals("123,456,789.13", fn.format(123456789.126, "#,###.00"));
    assertEquals("1,123,456,789.00", fn.format(1123456789, "#,###.00"));
  }

  @Test
  void roundNumber() {
    assertNull(fn.round(null, 0));
    assertNull(fn.round(null, 1));
    assertNull(fn.round(null, 2));
    assertEquals(new BigDecimal(123456789), fn.round(123456789, 0));
    assertEquals(new BigDecimal(123456789), fn.round(123456789.45, 0));
    assertEquals(new BigDecimal("123456789.5"), fn.round(123456789.45, 1));
    assertEquals(new BigDecimal("123456789.45"), fn.round(123456789.454, 2));
    assertEquals(new BigDecimal("123456789.46"), fn.round(123456789.455, 2));
    assertEquals(new BigDecimal("123456789.46"), fn.round(123456789.456, 2));
  }

  @Test
  void toInt() {
    assertNull(fn.toInt(null));
    assertEquals(123, fn.toInt("123"));
  }

  @Test
  void concatString() {
    assertNull(fn.concat());
    assertEquals("abc", fn.concat("abc"));
    assertEquals("abc", fn.concat("ab", "c"));
  }

  @Test
  void formatDateTime() {
    assertNull(fn.format((TemporalAccessor) null, null));
    assertEquals("2017", fn.format(Year.of(2017), "yyyy"));
    assertEquals("2017-01", fn.format(YearMonth.of(2017, 1), "yyyy-MM"));
    assertEquals("01", fn.format(Month.of(1), "MM"));
    assertEquals("2017-01-02", fn.format(LocalDate.of(2017, 1, 2), "yyyy-MM-dd"));
    assertEquals("2017/1/12", fn.format(LocalDate.of(2017, 1, 12), "yyyy/M/d"));
    assertEquals("2017-01-02 10:20:30", fn.format(LocalDateTime.of(2017, 1, 2, 10, 20, 30), "yyyy-MM-dd HH:mm:ss"));
    assertEquals("13:20:30", fn.format(LocalTime.of(13, 20, 30), "HH:mm:ss"));
  }

  @Test
  void join() {
    assertEquals("s1, s2", fn.join(Arrays.asList("s1", null, "s2")));
    assertEquals("s1-s2", fn.join(Arrays.asList("s1", null, "s2"), "-"));
  }

  @Test
  void joinProperty() {
    Map<String, Object> map = new HashMap<>();
    map.put("p1", "m1");
    Bean bean = new Bean();
    bean.setP1("b1");
    assertEquals("m1, b1", fn.joinProperty(Arrays.asList(map, bean), "p1"));
    assertEquals("m1,b1", fn.joinProperty(Arrays.asList(map, bean), "p1", ","));
    assertEquals("", fn.joinProperty(Arrays.asList(map, bean), "p2", ","));

    bean = new Bean();
    bean.setP2(2);
    assertEquals("2", fn.joinProperty(Arrays.asList(map, bean), "p2", ","));

    bean = new Bean();
    bean.setP3(3);
    assertEquals("3", fn.joinProperty(Arrays.asList(map, bean), "p3", ","));
  }

  @Test
  void timeSubtract() {
    assertNull(fn.timeSubtract( null, null));
    LocalDateTime time = LocalDateTime.now();
    assertEquals("10m",fn.timeSubtract(time,time.plusMinutes(10)));
    assertEquals("10h10m",fn.timeSubtract(time,time.plusMinutes(10).plusHours(10)));
    assertEquals("10d10h10m",fn.timeSubtract(time,time.plusMinutes(10).plusHours(10).plusDays(10)));
  }

  @Test
  void renderTemplate() throws Exception {
    // template
    InputStream template = getClass().getClassLoader().getResourceAsStream("templates/common-functions.xlsx");

    // output to
    File out = new File("target/common-functions-result.xlsx");
    if (out.exists()) out.delete();
    OutputStream output = new FileOutputStream(out);

    // data
    Context context = new Context();

    //-- inject common-functions
    context.putVar("fn", CommonFunctions.getSingleton());

    //-- other data
    context.putVar("num", new BigDecimal("123.456"));
    context.putVar("datetime", LocalDateTime.now());
    context.putVar("date", LocalDate.now());
    context.putVar("time", LocalTime.now());
    context.putVar("str", "123");
    context.putVar("beans", Arrays.asList(new Bean("p1"), new Bean("p2")));
    context.putVar("strings", Arrays.asList("s1", "s2"));

    // render
    JxlsHelper.getInstance().processTemplate(template, output, context);
  }
}
package tech.simter.jxls.ext;

import org.junit.Test;
import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.time.*;
import java.time.temporal.TemporalAccessor;

import static org.hamcrest.CoreMatchers.is;
import static org.hamcrest.CoreMatchers.nullValue;
import static org.junit.Assert.assertThat;

/**
 * Common Functions test.
 *
 * @author RJ
 */
public class CommonFunctionsTest {
  private final CommonFunctions fn = CommonFunctions.getSingleton();

  @Test
  public void formatNumber() {
    assertThat(fn.round(null, 0), nullValue());
    assertThat(fn.format(123456789.1, "#,###.00"), is("123,456,789.10"));
    assertThat(fn.format(123456789.12, "#,###.00"), is("123,456,789.12"));
    assertThat(fn.format(123456789.124, "#,###.00"), is("123,456,789.12"));
    assertThat(fn.format(123456789.125, "#,###.00"), is("123,456,789.12"));
    assertThat(fn.format(123456789.126, "#,###.00"), is("123,456,789.13"));
    assertThat(fn.format(1123456789, "#,###.00"), is("1,123,456,789.00"));
  }

  @Test
  public void roundNumber() {
    assertThat(fn.round(null, 0), nullValue());
    assertThat(fn.round(null, 1), nullValue());
    assertThat(fn.round(null, 2), nullValue());
    assertThat(fn.round(123456789, 0), is(new BigDecimal(123456789)));
    assertThat(fn.round(123456789.45, 0), is(new BigDecimal(123456789)));
    assertThat(fn.round(123456789.45, 1), is(new BigDecimal("123456789.5")));
    assertThat(fn.round(123456789.454, 2), is(new BigDecimal("123456789.45")));
    assertThat(fn.round(123456789.455, 2), is(new BigDecimal("123456789.46")));
    assertThat(fn.round(123456789.456, 2), is(new BigDecimal("123456789.46")));
  }

  @Test
  public void toInt() {
    assertThat(fn.toInt(null), nullValue());
    assertThat(fn.toInt("123"), is(123));
  }

  @Test
  public void concatString() {
    assertThat(fn.concat(), nullValue());
    assertThat(fn.concat("abc"), is("abc"));
    assertThat(fn.concat("ab", "c"), is("abc"));
  }

  @Test
  public void formatDateTime() {
    assertThat(fn.format((TemporalAccessor) null, null), nullValue());
    assertThat(fn.format(Year.of(2017), "yyyy"), is("2017"));
    assertThat(fn.format(YearMonth.of(2017, 1), "yyyy-MM"), is("2017-01"));
    assertThat(fn.format(Month.of(1), "MM"), is("01"));
    assertThat(fn.format(LocalDate.of(2017, 1, 2), "yyyy-MM-dd"), is("2017-01-02"));
    assertThat(fn.format(LocalDate.of(2017, 1, 12), "yyyy/M/d"), is("2017/1/12"));
    assertThat(fn.format(LocalDateTime.of(2017, 1, 2, 10, 20, 30), "yyyy-MM-dd HH:mm:ss"), is("2017-01-02 10:20:30"));
    assertThat(fn.format(LocalTime.of(13, 20, 30), "HH:mm:ss"), is("13:20:30"));
  }

  @Test
  public void renderTemplate() throws Exception {
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

    // render
    JxlsHelper.getInstance().processTemplate(template, output, context);
  }
}
package tech.simter.jxls.ext;

import org.junit.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.time.LocalDateTime;
import java.time.YearMonth;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static org.hamcrest.CoreMatchers.is;
import static org.junit.Assert.assertThat;

/**
 * The Excel Utils test.
 *
 * @author RJ
 */
public class JxlsUtilsTest {
  @Test
  public void xlsx() throws Exception {
    // template
    InputStream template = getClass().getClassLoader().getResourceAsStream("templates/common-functions.xlsx");

    // output to
    File out = new File("target/common-functions-result.xlsx");
    if (out.exists()) out.delete();
    OutputStream output = new FileOutputStream(out);

    // template data
    Map<String, Object> data = generateData();

    // render
    JxlsUtils.renderTemplate(template, data, output);

    // verify
    assertThat(out.exists(), is(true));
    assertThat(out.getTotalSpace() > 0, is(true));
  }

  @Test
  public void xls() throws Exception {
    // from template
    InputStream template = getClass().getClassLoader().getResourceAsStream("templates/common-functions.xls");

    // output to
    File out = new File("target/common-functions-result.xls");
    if (out.exists()) out.delete();
    OutputStream output = new FileOutputStream(out);

    // template data
    Map<String, Object> data = generateData();

    // render
    JxlsUtils.renderTemplate(template, data, output);

    // verify
    assertThat(out.exists(), is(true));
    assertThat(out.getTotalSpace() > 0, is(true));
  }

  // generate test data
  private Map<String, Object> generateData() {
    Map<String, Object> data = new HashMap<>();
    data.put("stage", 0);
    Map<Integer, Object> stageLabels = new HashMap<>();
    stageLabels.put(0, "TODO");
    stageLabels.put(1, "ALLOW");
    data.put("stageLabels", stageLabels);

    List<Map<String, Object>> rows = new ArrayList<>();
    data.put("rows", rows);
    Map<String, Object> row;
    for (int i = 0; i < 10; i++) {
      row = new HashMap<>();
      rows.add(row);
      row.put("yearMonth", YearMonth.now());
      row.put("dateTime", LocalDateTime.now());
      row.put("str", "test");
      row.put("money", new BigDecimal("100.01").add(new BigDecimal((float) i / 1000)).setScale(2, BigDecimal.ROUND_HALF_UP));
      row.put("stage", i % 2);
      if (i % 2 == 0) {
        row.put("remark", "remark" + (i + 1));
      }
    }
    return data;
  }
}
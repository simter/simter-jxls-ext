package tech.simter.jxls;

import org.junit.jupiter.api.Test;
import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.time.OffsetDateTime;
import java.time.temporal.ChronoUnit;
import java.util.ArrayList;
import java.util.List;

import static java.time.temporal.ChronoUnit.SECONDS;
import static org.assertj.core.api.Assertions.assertThat;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * Excel charts with fixed size collection test.
 *
 * @author RJ
 */
public class ChartsTest {
  @SuppressWarnings("ResultOfMethodCallIgnored")
  @Test
  public void test() throws Exception {
    // template
    InputStream template = getClass().getClassLoader().getResourceAsStream("templates/charts.xlsx");

    // output to
    File out = new File("target/charts-result.xlsx");
    if (out.exists()) out.delete();
    OutputStream output = new FileOutputStream(out);

    // template data
    Context context = new Context();
    List<Item> rows = new ArrayList<>();
    rows.add(new Item("Derek", 3000, 2000));
    rows.add(new Item("Elsa", 1500, 500));
    rows.add(new Item("Oleg", 2300, 1300));
    rows.add(new Item("Neil", 2500,1500));
    rows.add(new Item("Maria", 1700, 700));
    rows.add(new Item("John", 2800, 2000));
    rows.add(new Item("Leonid", 1700, 1000));
    context.putVar("items", rows);
    context.putVar("title", "X-Y");
    context.putVar("ts", OffsetDateTime.now().truncatedTo(SECONDS).toString());

    // render
    // 必须设置 .setEvaluateFormulas(true)，否则生成的 Excel 文件，
    // 用到公式的地方不会显示公式的结果，需要双击单元格回车才能看到公式结果。
    JxlsHelper.getInstance()
      .setEvaluateFormulas(true)
      .processTemplate(template, output, context);

    // verify
    assertTrue(out.exists());
    assertThat(out.getTotalSpace()).isGreaterThan(0);
  }

  public static class Item {
    private final String x;
    private final int y1;
    private final int y2;

    public Item(String x, int y1, int y2) {
      this.x = x;
      this.y1 = y1;
      this.y2 = y2;
    }

    public String getX() {
      return x;
    }

    public int getY1() {
      return y1;
    }

    public int getY2() {
      return y2;
    }
  }
}
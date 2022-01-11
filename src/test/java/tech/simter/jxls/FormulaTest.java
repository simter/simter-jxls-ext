package tech.simter.jxls;

import org.junit.jupiter.api.Test;
import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * Excel formula test.
 *
 * @author RJ
 */
public class FormulaTest {
  @Test
  public void test() throws Exception {
    // template
    InputStream template = getClass().getClassLoader().getResourceAsStream("templates/formula.xlsx");

    // output to
    File out = new File("target/formula-result.xlsx");
    if (out.exists()) out.delete();
    OutputStream output = new FileOutputStream(out);

    // template data
    Context context = new Context();
    List<Integer> rows = new ArrayList<>();
    rows.add(1);
    rows.add(2);
    rows.add(3);
    context.putVar("rows", rows);
    context.putVar("title", "测试");

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
}
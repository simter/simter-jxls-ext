package tech.simter.jxls.ext;

import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * The Jxls test.
 *
 * @author RJ
 */
class TwoSubListTest {
  @Test
  void test() throws Exception {
    // template
    InputStream template = getClass().getClassLoader().getResourceAsStream("templates/two-sub-list.xlsx");

    // output to
    File out = new File("target/two-sub-list-result.xlsx");
    if (out.exists()) out.delete();
    OutputStream output = new FileOutputStream(out);

    // template data
    Map<String, Object> data = generateData();

    // render
    JxlsUtils.renderTemplate(template, data, output);

    // verify
    assertThat(out.getTotalSpace()).isGreaterThan(0);
  }

  private Map<String, Object> generateData() {
    Map<String, Object> data = new HashMap<>();
    data.put("subject", "JXLS two-sub-list test");

    List<Map<String, Object>> rows = new ArrayList<>();
    data.put("rows", rows);
    int rowNumber = 0;
    boolean align = true;
    rows.add(createRow(++rowNumber, 2, 2, align));
    rows.add(createRow(++rowNumber, 2, 1, align));
    rows.add(createRow(++rowNumber, 2, 0, align)); // 0 代表空的集合（不是 null 而是 size = 0）

    // -1 代表不存在的集合，会导致 jx:each 报错：org.jxls.common.JxlsException: r.subs2 expression is not a collection
    //rows.add(createRow(++rowNumber, 2, -1, align));

    rows.add(createRow(++rowNumber, 1, 1, align));
    rows.add(createRow(++rowNumber, 1, 0, align));
    rows.add(createRow(++rowNumber, 0, 0, align));

    rows.add(createRow(++rowNumber, 1, 2, align));
    rows.add(createRow(++rowNumber, 0, 2, align));

    return data;
  }

  @SuppressWarnings("unchecked")
  private Map<String, Object> createRow(int rowNumber, int leftSubsCount, int rightSubsCount, boolean align) {
    Map<String, Object> row = new HashMap<>();
    row.put("sn", rowNumber);
    row.put("name", "row" + rowNumber);
    if (leftSubsCount >= 0) row.put("subs1", createSubs(rowNumber, 1, leftSubsCount));
    if (rightSubsCount >= 0) row.put("subs2", createSubs(rowNumber, 2, rightSubsCount));

    // 如果 leftSubsCount != rightSubsCount, 取两者的最大值，将各自的集合追加空元素来对齐集合的总长度
    // 目的是保证不会生成没有格式的空白单元格
    List<Map<String, Object>> subs;
    if (align && (leftSubsCount != rightSubsCount || leftSubsCount == 0)) {
      int maxCount = Math.max(1, Math.max(leftSubsCount, rightSubsCount));
      if (leftSubsCount < maxCount) {
        subs = (List<Map<String, Object>>) row.get("subs1");
        for (int i = 0; i < maxCount - leftSubsCount; i++) subs.add(new HashMap<>());
      }
      if (rightSubsCount < maxCount) {
        subs = (List<Map<String, Object>>) row.get("subs2");
        for (int i = 0; i < maxCount - rightSubsCount; i++) subs.add(new HashMap<>());
      }
    }
    return row;
  }

  private List<Map<String, Object>> createSubs(int rowNumber, int subNumber, int count) {
    List<Map<String, Object>> subs = new ArrayList<>();
    Map<String, Object> sub;
    for (int i = 1; i <= count; i++) {
      sub = new HashMap<>();
      subs.add(sub);
      sub.put("sn", rowNumber + "-" + i);
      sub.put("name", "row" + rowNumber + "sub" + subNumber + "-" + i);
    }
    return subs;
  }
}
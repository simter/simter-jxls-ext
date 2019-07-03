package tech.simter.jxls.ext;

import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.*;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * The Jxls test.
 *
 * @author RJ
 */
class DynamicColumnTest {
  @Test
  void test() throws Exception {
    // template
    InputStream template = getClass().getClassLoader().getResourceAsStream("templates/dynamic-column.xlsx");

    // output to
    File out = new File("target/dynamic-column-result.xlsx");
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
    data.put("subject", "JXLS dynamic columns test");

    List<Map<String, Object>> rows = new ArrayList<>();
    data.put("rows", rows);
    List<String> itemNames = Arrays.asList("Item1", "Item2", "Item3");
    data.put("itemNames", itemNames);
    int rowNumber = 0;
    rows.add(createRow(++rowNumber, itemNames));
    rows.add(createRow(++rowNumber, itemNames));
    rows.add(createRow(++rowNumber, itemNames));

    return data;
  }

  private Map<String, Object> createRow(int rowNumber, List<String> itemNames) {
    Map<String, Object> row = new HashMap<>();
    row.put("sn", rowNumber);
    row.put("name", "row" + rowNumber);

    List<Object> itemValues = new ArrayList<>();
    int i = 0;
    for (String ignored : itemNames) {
      itemValues.add(100000000.345678 * rowNumber + (++i));
    }
    row.put("itemValues", itemValues);

    return row;
  }
}
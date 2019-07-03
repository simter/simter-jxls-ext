package tech.simter.jxls.ext;

import org.junit.jupiter.api.Test;
import org.jxls.builder.xls.XlsCommentAreaBuilder;
import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;
import static tech.simter.jxls.ext.JxlsUtils.convert2Context;

/**
 * The each-merge command test.
 *
 * @author RJ
 */
class EachMergeCommandTest {
  // one main command with one sub command
  @Test
  void mergeWithOneSubCommand() throws Exception {
    // global add custom each-merge command to XlsCommentAreaBuilder
    XlsCommentAreaBuilder.addCommandMapping(EachMergeCommand.COMMAND_NAME, EachMergeCommand.class);

    // template
    InputStream template = getClass().getClassLoader().getResourceAsStream("templates/each-merge.xlsx");

    // output to
    File out = new File("target/each-merge-result.xlsx");
    if (out.exists()) out.delete();
    OutputStream output = new FileOutputStream(out);

    // generate template data
    Context context = convert2Context(generateData());

    // render
    JxlsHelper.getInstance().processTemplate(template, output, context);

    // verify
    assertThat(out.getTotalSpace()).isGreaterThan(0);
  }

  @Test
  void mergeWithOneSubCommand1() throws Exception {
    // template
    InputStream template = getClass().getClassLoader().getResourceAsStream("templates/each-merge.xlsx");

    // output to
    File out = new File("target/each-merge-result.xlsx");
    if (out.exists()) out.delete();
    OutputStream output = new FileOutputStream(out);

    // generate template data
    Map<String, Object> data = generateData();

    // render
    JxlsUtils.renderTemplate(template, data, output);

    // verify
    assertThat(out.getTotalSpace()).isGreaterThan(0);
  }

  // one main command with two sub commands
  @Test
  void mergeWithTwoSubCommand() throws Exception {
    // template
    InputStream template = getClass().getClassLoader().getResourceAsStream("templates/each-merge2.xlsx");

    // output to
    File out = new File("target/each-merge2-result.xlsx");
    if (out.exists()) out.delete();
    OutputStream output = new FileOutputStream(out);

    // generate template data
    Map<String, Object> data = generateData();
    copySubsToSubs1(data); // create the second sub command data

    // render
    JxlsUtils.renderTemplate(template, data, output);

    // verify
    assertThat(out.getTotalSpace()).isGreaterThan(0);
  }

  @SuppressWarnings("unchecked")
  private void copySubsToSubs1(Map<String, Object> data) {
    ((List<Map<String, Object>>) data.get("rows")).forEach(row -> {
      List<Map<String, Object>> subs = (List<Map<String, Object>>) row.get("subs");
      List<Map<String, Object>> subs1 = new ArrayList<>();
      row.put("subs1", subs1);
      subs.forEach(sub -> {
        HashMap<String, Object> newMap = new HashMap<>(sub);
        for (Map.Entry<String, Object> e : newMap.entrySet()) {
          e.setValue(e.getValue() + "-s");
        }
        subs1.add(newMap);
      });
    });
  }

  private static Map<String, Object> generateData() {
    Map<String, Object> data = new HashMap<>();
    data.put("subject", "JXLS merge cell test");

    List<Map<String, Object>> rows = new ArrayList<>();
    data.put("rows", rows);
    int rowNumber = 0;
    //rows.add(createRow(++rowNumber, 3));
    rows.add(createRow(++rowNumber, 3));
    rows.add(createRow(++rowNumber, 1));
    rows.add(createRow(++rowNumber, 2));
    //rows.add(createRow(++rowNumber, 0));

    return data;
  }

  private static Map<String, Object> createRow(int rowNumber, int subsCount) {
    Map<String, Object> row = new HashMap<>();
    row.put("sn", rowNumber);
    row.put("name", "row" + rowNumber);
    if (subsCount >= 0) row.put("subs", createSubs(rowNumber, subsCount));
    return row;
  }

  private static List<Map<String, Object>> createSubs(int rowNumber, int count) {
    List<Map<String, Object>> subs = new ArrayList<>();
    Map<String, Object> sub;
    for (int i = 1; i <= count; i++) {
      sub = new HashMap<>();
      subs.add(sub);
      sub.put("sn", rowNumber + "-" + i);
      sub.put("name", "row" + rowNumber + "sub" + +i);
    }
    return subs;
  }
}
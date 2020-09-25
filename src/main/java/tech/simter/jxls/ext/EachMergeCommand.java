package tech.simter.jxls.ext;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.jxls.area.Area;
import org.jxls.command.CellRefGenerator;
import org.jxls.command.EachCommand;
import org.jxls.common.*;
import org.jxls.transform.Transformer;
import org.jxls.transform.poi.PoiTransformer;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.ArrayList;
import java.util.Collection;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

/**
 * Extends {@link EachCommand} to support merge cells.
 *
 * @author RJ
 */
public class EachMergeCommand extends EachCommand {
  private static final Logger logger = LoggerFactory.getLogger(EachMergeCommand.class);

  public static final String COMMAND_NAME = "each-merge";

  public EachMergeCommand() {
    super();
  }

  public EachMergeCommand(String var, String items, Direction direction) {
    super(var, items, direction);
  }

  public EachMergeCommand(String items, Area area) {
    super(items, area);
  }

  public EachMergeCommand(String var, String items, Area area) {
    super(var, items, area);
  }

  public EachMergeCommand(String var, String items, Area area, Direction direction) {
    super(var, items, area, direction);
  }

  public EachMergeCommand(String var, String items, Area area, CellRefGenerator cellRefGenerator) {
    super(var, items, area, cellRefGenerator);
  }

  @Override
  public Size applyAt(CellRef cellRef, Context context) {
    // collect sub command areas
    List<Area> childAreas = this.getAreaList().stream()
      .flatMap(area1 -> area1.getCommandDataList().stream())
      .flatMap(commandData -> commandData.getCommand().getAreaList().stream())
      .collect(Collectors.toList());
    List<AreaRef> childAreaRefs = childAreas.stream()
      .map(Area::getAreaRef).collect(Collectors.toList());

    // register AreaListener for parent command area
    Area parentArea = this.getAreaList().get(0);
    MergeCellListener listener = new MergeCellListener(getTransformer(), parentArea.getAreaRef(), childAreaRefs,
      ((Collection<?>) context.getVar(this.getItems())).size()); // 总数据量
    logger.info("register MergeCellListener@{} for parent-area '{}', cellRef={}", listener.hashCode(), parentArea.getAreaRef(), cellRef);
    parentArea.addAreaListener(listener);

    // register AreaListener for all sub command area
    childAreas.forEach(area -> {
      logger.info("register MergeCellListener@{} for child-area '{}', cellRef={}, parent={}", listener.hashCode(), area.getAreaRef(), cellRef, parentArea.getAreaRef());
      area.addAreaListener(listener);
    });

    // standard dealing
    return super.applyAt(cellRef, context);
  }

  /**
   * The {@link AreaListener} for merge cells.
   */
  public static class MergeCellListener implements AreaListener {
    private final PoiTransformer transformer;
    private final int parentStartColumn;                   // parent command start column
    private final int[] childStartColumns;                 // all sub command start column
    private final int[] mergeColumns;                      // to merge columns
    private final List<int[]> records = new ArrayList<>(); // 0 - start column, 1 - end column
    private final int parentCount;

    private int childRow;
    private int parentProcessed;
    private String sheetName;

    MergeCellListener(Transformer transformer, AreaRef parent, List<AreaRef> children, int parentCount) {
      this.transformer = (PoiTransformer) transformer;
      this.parentCount = parentCount;
      this.parentStartColumn = parent.getFirstCellRef().getCol();

      // find all sub command columns
      int[] childCols = children.stream()
        .flatMapToInt(ref -> IntStream.range(ref.getFirstCellRef().getCol(), ref.getLastCellRef().getCol() + 1))
        .distinct().sorted()
        .toArray();

      // find all sub command start column
      this.childStartColumns = children.stream()
        .mapToInt(ref -> ref.getFirstCellRef().getCol())
        .distinct().sorted().toArray();

      // get columns to merge by filter childCols
      this.mergeColumns = IntStream.range(parent.getFirstCellRef().getCol(), parent.getLastCellRef().getCol() + 1)
        .filter(parentCol -> IntStream.of(childCols).noneMatch(childCol -> childCol == parentCol))
        .toArray();

      if (logger.isDebugEnabled()) {
        logger.debug("parentArea={}", parent);
        logger.debug("parentStartColumn={}", parentStartColumn);
        logger.debug("childStartColumns={}", childStartColumns);
        logger.debug("mergeColumns={}", mergeColumns);
        logger.debug("childCols={}", childCols);
      }
    }

    @Override
    public void beforeApplyAtCell(CellRef cellRef, Context context) {
    }

    @Override
    public void afterApplyAtCell(CellRef cellRef, Context context) {
    }

    @Override
    public void beforeTransformCell(CellRef srcCell, CellRef targetCell, Context context) {
    }

    // remember that all sub command AreaListener is invoked before parent command AreaListener.
    // This class use this feature to do the merge work.
    @Override
    public void afterTransformCell(CellRef srcCell, CellRef targetCell, Context context) {
      if (parentProcessed == 0) this.sheetName = targetCell.getSheetName();

      if (targetCell.getCol() == parentStartColumn) { // main command process
        this.parentProcessed++;

        logger.debug("parent: srcCell={}, targetCell={} [{}, {}]", srcCell, targetCell,
          targetCell.getRow(), targetCell.getCol());

        //should be recorded just on necessary
        if (targetCell.getRow() < this.childRow) this.records.add(new int[]{targetCell.getRow(), this.childRow});

        // merge work invoke on the last
        if (this.parentProcessed == this.parentCount) {
          Workbook workbook = transformer.getWorkbook();
          Sheet sheet = workbook.getSheet(sheetName);
          doMerge(sheet, this.records, this.mergeColumns, srcCell);
        }
        this.childRow = 0;

        // record the current row number of sub command process
      } else if (IntStream.of(childStartColumns).anyMatch(col -> col == targetCell.getCol())) {
        this.childRow = Math.max(this.childRow, targetCell.getRow());

        logger.debug("child: srcCell={}, targetCell={} [{}, {}]", srcCell, targetCell,
          targetCell.getRow(), targetCell.getCol());
      }
    }

    private static void doMerge(Sheet sheet, List<int[]> records, int[] mergeColumns, CellRef srcCell) {
      if (logger.isDebugEnabled()) {
        logger.debug("start merge: sheetName={}, records={}", sheet.getSheetName(),
          records.stream().map(startEnd -> "[" + startEnd[0] + "," + startEnd[1] + "]")
            .collect(Collectors.joining(",")));
      }
      records.forEach(startEnd -> merge4Row(sheet, startEnd[0], startEnd[1], mergeColumns, srcCell));
    }

    private static void merge4Row(Sheet sheet, int fromRow, int toRow, int[] mergeColumns, CellRef srcCell) {
      if (fromRow >= toRow) {
        logger.warn("  No need to merge because same row：fromRow={}, toRow={}", fromRow, toRow);
        return;
      }
      Cell originCell;
      CellStyle originCellStyle;
      CellRangeAddress region;
      for (int col : mergeColumns) {
        logger.debug("  fromRow={}, toRow={}, col={}", fromRow, toRow, col);
        region = new CellRangeAddress(fromRow, toRow, col, col);
        sheet.addMergedRegion(region);

        //firstCell = sheet.getRow(fromRow).getCell(col);
        originCell = sheet.getRow(srcCell.getRow()).getCell(col);
        if (originCell == null) {
          logger.info("  Missing cell: row={}, col={}", fromRow, col);
        }
        if (originCell != null) {
          // copy originCell style to the merged cell
          originCellStyle = originCell.getCellStyle();
          RegionUtil.setBorderTop(originCellStyle.getBorderTop(), region, sheet);
          RegionUtil.setBorderRight(originCellStyle.getBorderRight(), region, sheet);
          RegionUtil.setBorderBottom(originCellStyle.getBorderBottom(), region, sheet);
          RegionUtil.setBorderLeft(originCellStyle.getBorderLeft(), region, sheet);
        } else {
          RegionUtil.setBorderTop(BorderStyle.THIN, region, sheet);
          RegionUtil.setBorderRight(BorderStyle.THIN, region, sheet);
          RegionUtil.setBorderBottom(BorderStyle.THIN, region, sheet);
          RegionUtil.setBorderLeft(BorderStyle.THIN, region, sheet);
        }
      }
    }
  }
}
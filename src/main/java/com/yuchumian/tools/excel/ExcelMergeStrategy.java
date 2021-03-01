package com.yuchumian.tools.excel;

import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.write.merge.AbstractMergeStrategy;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.*;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

/**
 * @author yuchumian 2021-02-26
 **/
@Slf4j
public class ExcelMergeStrategy extends AbstractMergeStrategy {

    /**
     * 需要合并的列编号，从0开始
     */
    private Set<Integer> mergeColumnIndex = new HashSet<>();

    /**
     * 数据集大小，用于区别结束行位置
     */
    private Integer maxRow = 0;

    /**
     * 禁止无参声明
     */
    private ExcelMergeStrategy() {
        log.error("not support!");
    }

    public ExcelMergeStrategy(Integer maxRow, int... mergeColumnIndex) {
        this.mergeColumnIndex = IntStream.of(mergeColumnIndex).boxed().collect(Collectors.toCollection(HashSet::new));
        this.maxRow = maxRow;
    }

    /**
     * 记录上一次合并行的信息
     */
    private final Map<Integer, MergeRange> recordRow = new HashMap<>();

    /**
     * merge
     */
    @Override
    protected void merge(Sheet sheet, Cell cell, Head head, Integer relativeRowIndex) {
        int currentCellIndex = cell.getColumnIndex();
        // 判断该列是否需要合并
        if (mergeColumnIndex.contains(currentCellIndex)) {
            String currentCellValue = cell.getStringCellValue();
            int currentRowIndex = cell.getRowIndex();
            if (!recordRow.containsKey(currentCellIndex)) {
                // 记录起始位置
                recordRow.put(currentCellIndex, new MergeRange(currentCellValue, currentRowIndex, currentRowIndex, currentCellIndex, currentCellIndex));
                return;
            }
            // 用上一行这列的值做对比.
            MergeRange mergeRange = recordRow.get(currentCellIndex);
            if (!(mergeRange.lastValue != null && mergeRange.lastValue.equals(currentCellValue))) {
                // 结束的位置触发下合并.
                // 同行同列不能合并，会抛异常
                if (mergeRange.startRow != mergeRange.endRow || mergeRange.startCell != mergeRange.endCell) {
                    sheet.addMergedRegionUnsafe(new CellRangeAddress(mergeRange.startRow, mergeRange.endRow, mergeRange.startCell, mergeRange.endCell));
                }
                // 更新当前列起始位置
                recordRow.put(currentCellIndex, new MergeRange(currentCellValue, currentRowIndex, currentRowIndex, currentCellIndex, currentCellIndex));
            }
            // 合并行 + 1
            mergeRange.endRow += 1;
            // 结束的位置触发下最后一次没完成的合并
            if (relativeRowIndex.equals(maxRow - 1)) {
                MergeRange lastMergeRange = recordRow.get(currentCellIndex);
                // 同行同列不能合并，会抛异常
                if (lastMergeRange.startRow != lastMergeRange.endRow || lastMergeRange.startCell != lastMergeRange.endCell) {
                    sheet.addMergedRegionUnsafe(new CellRangeAddress(lastMergeRange.startRow, lastMergeRange.endRow, lastMergeRange.startCell, lastMergeRange.endCell));
                }
            }
        }
    }
}

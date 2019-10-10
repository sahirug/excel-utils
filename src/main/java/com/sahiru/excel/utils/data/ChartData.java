package com.sahiru.excel.utils.data;

import java.util.HashMap;

public class ChartData extends HashMap<String, Object> {
    private int labelColumnIndex, valueColumnIndex;

    public ChartData(int labelColumnIndex, int valueColumnIndex) {
        this.labelColumnIndex = labelColumnIndex;
        this.valueColumnIndex = valueColumnIndex;
    }

    public int getLabelColumnIndex() {
        return labelColumnIndex;
    }

    public int getValueColumnIndex() {
        return valueColumnIndex;
    }
}

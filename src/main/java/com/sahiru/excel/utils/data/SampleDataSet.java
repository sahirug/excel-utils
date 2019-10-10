package com.sahiru.excel.utils.data;

import java.util.ArrayList;
import java.util.List;
import java.util.Random;

public class SampleDataSet {
    public static List<ChartData> SAMPLE_DATA = new ArrayList<ChartData>();
    private static final int START_ROW = 7;
    private static final int DATA_SIZE = 3;
    private static final int NUMBER_OF_CHARTS = 3;
    private static final String[] LABELS = {"Release1", "Release2", "Release3"};
    private static final Random RANDOM = new Random();


    static {
        for (int i = START_ROW; i < (START_ROW+DATA_SIZE); i++) {
            int[] columnIndexes = getColumnIndexes(i);
            ChartData data = new ChartData(columnIndexes[0], columnIndexes[1]);
            for (int j = 0; j < LABELS.length; j++) {
                data.put(LABELS[j], RANDOM.nextInt(50));
            }
            SAMPLE_DATA.add(data);
        }
    }

    private static int[] getColumnIndexes(int j) {
        switch (j) {
            case 7:
                return new int[]{0,1};
            case 8:
                return new int[]{11,14};
            case 9:
                return new int[]{28,31};
            default:
                return new int[]{};
        }
    }
}

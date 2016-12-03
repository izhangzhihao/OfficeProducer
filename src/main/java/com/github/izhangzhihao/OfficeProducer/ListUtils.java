package com.github.izhangzhihao.OfficeProducer;


import java.util.List;

public class ListUtils {
    /**
     * 判断一个List是否是NULL或者是空
     *
     * @param list 要判断的List
     * @return 结果
     */
    public static boolean isNullOrEmpty(final List list) {
        return list == null || list.isEmpty();
    }
}

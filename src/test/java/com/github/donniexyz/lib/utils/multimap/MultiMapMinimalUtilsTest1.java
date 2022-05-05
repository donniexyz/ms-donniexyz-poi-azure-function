package com.github.donniexyz.lib.utils.multimap;

import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

import java.util.HashMap;
import java.util.Map;

class MultiMapMinimalUtilsTest1 {

    @Test
    void getFromMultiMap() {
        Map<String, Object> map = new HashMap<>();

        map.put("1", "1V");

        Assertions.assertEquals("1V", MultiMapMinimalUtils.getFromMultiMap("1", map, ""));
    }
}
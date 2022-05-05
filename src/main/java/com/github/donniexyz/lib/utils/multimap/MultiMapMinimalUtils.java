package com.github.donniexyz.lib.utils.multimap;

import java.util.Map;

public class MultiMapMinimalUtils {


    public static <T> T defaultIfNull(T object, T defaultValue) {
        return null == object ? defaultValue : object;
    }

    /**
     * this is to get value from Map<String, Map<String, Map<String ...>>>
     * for example: key = "dataVO.customers.address.line1" --> (String) map.get("dataVO").get("customers").get("address").get("line1")
     *
     * @param key          if null then will return the map
     * @param map
     * @param defaultValue
     * @return
     */
    public static Object getFromMultiMap(String key, Map<String, Object> map, Object defaultValue) {
        try {
            if (null == map) return defaultValue;
            if (null == key) return map.getOrDefault(null, defaultValue);
            String[] k = key.split("\\.");
            Map<String, Object> mapIter = map;
            for (int i = 0; i < k.length - 1; i++) {
                if (null == mapIter)
                    return defaultValue;
                mapIter = (Map<String, Object>) mapIter.get(k[i]);
            }
            if (null == mapIter)
                return defaultValue;
            String lastKey = k[k.length - 1];
            return mapIter.getOrDefault(lastKey, defaultValue);
        } catch (RuntimeException e) {
            return defaultValue;
        }
    }

}

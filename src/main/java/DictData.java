import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class DictData {
        public static Map<String, Object> dict = null;
        static {
            dict = new HashMap<>();
            List<String> list = new ArrayList<>();
            list.add("会议");
            list.add("公文");
            list.add("项目");
            list.add("测试大类");
            dict.put("car-dict", list);

            Map<String, List<String>> subMap = new HashMap<>();
            list = new ArrayList<>();
            list.add("2018");
            list.add("一般");
            list.add("重点");
            subMap.put("会议", list);
            list = new ArrayList<>();
            list.add("党政办公文");
            subMap.put("公文", list);
            list = new ArrayList<>();
            subMap.put("项目", list);
            list = new ArrayList<>();
            subMap.put("测试大类", list);
            dict.put("fruit-dict", subMap);

            list = new ArrayList<>();
            list.add("汽车-1");
            list.add("水果-1");
            dict.put("t-dict", list);
        }

        /** 获取数据字典中的值 */
        public static Object getDict(String dict) {
            return DictData.dict.get(dict);
        }
}

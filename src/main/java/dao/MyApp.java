package dao;

import org.nutz.ioc.Ioc;
import org.nutz.ioc.impl.NutIoc;
import org.nutz.ioc.loader.combo.ComboIocLoader;

public class MyApp {
    public static Ioc ioc;

    static {
        try {
            ioc = new NutIoc(new ComboIocLoader("json"));
        } catch (ClassNotFoundException e) {
            e.printStackTrace();
        }
    }
}

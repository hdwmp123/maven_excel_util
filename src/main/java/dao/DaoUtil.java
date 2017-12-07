package dao;

import org.nutz.dao.Dao;
import org.nutz.ioc.Ioc;
import org.nutz.ioc.impl.NutIoc;
import org.nutz.ioc.loader.json.JsonLoader;

import redis.clients.jedis.JedisPool;

public class DaoUtil {

    public DaoUtil() {
    }

    static Ioc ioc = new NutIoc(new JsonLoader(new String[] { "ioc.js" }));

    // static DataSource ds = ioc.get(javax.sql.DataSource.class);
    // static Dao dao = new NutDao(ds);
    public static Dao getDao() {
        // ioc.depose();
        // return dao;
        return ioc.get(Dao.class);
    }

    public static JedisPool get() {
        return ioc.get(JedisPool.class);
    }
}

package dao;

import org.nutz.dao.Dao;
import org.nutz.ioc.Ioc;
import org.nutz.ioc.impl.NutIoc;
import org.nutz.ioc.loader.json.JsonLoader;

import redis.clients.jedis.JedisPool;

/**
 * @author kingtiger
 */
public class DaoUtil {

    public DaoUtil() {
    }

    static Ioc ioc = new NutIoc(new JsonLoader(new String[]{"ioc.js"}));

    public static Dao getDao() {
        return ioc.get(Dao.class);
    }

    public static JedisPool get() {
        return ioc.get(JedisPool.class);
    }
}

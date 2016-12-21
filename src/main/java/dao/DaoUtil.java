package dao;

import javax.sql.DataSource;
import org.nutz.dao.Dao;
import org.nutz.dao.impl.NutDao;
import org.nutz.ioc.Ioc;
import org.nutz.ioc.impl.NutIoc;
import org.nutz.ioc.loader.json.JsonLoader;

public class DaoUtil {

	public DaoUtil() {
	}

	public static Dao getDao() {
		Ioc ioc = new NutIoc(new JsonLoader(new String[] { "ioc.js" }));
		DataSource ds = ioc.get(javax.sql.DataSource.class);
		Dao dao = new NutDao(ds);
		ioc.depose();
		return dao;
	}
}

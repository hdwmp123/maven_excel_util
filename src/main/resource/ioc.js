var ioc = {
	conf : {
		type : "org.nutz.ioc.impl.PropertiesProxy",
		fields : {
			paths : [ "config.properties" ]
		}
	},
	dataSource : {
		type : "org.nutz.dao.impl.SimpleDataSource",
		fields : {
			jdbcUrl : {
				java : "$conf.get('jdbcUrl')"
			},
			username : {
				java : "$conf.get('user')"
			},
			password : {
				java : "$conf.get('password')"
			},
		}
	},
	dao : {
		type : "org.nutz.dao.impl.NutDao",
		args : [ {
			refer : "dataSource"
		} ]
	},

	jedisPoolConfig : {
		type : "redis.clients.jedis.JedisPoolConfig",
		fields : {
			testWhileIdle : true,
			maxTotal : 100
		}
	},
	jedisPool : {
		type : "redis.clients.jedis.JedisPool",
		args : [ {
			refer : "jedisPoolConfig"
		}, {
			java : "$conf.get('redis.host', 'localhost')"
		}, {
			java : "$conf.getInt('redis.port', 6379)"
		}, {
			java : "$conf.getInt('redis.timeout', 2000)"
		}, {
			java : "$conf.get('redis.password')"
		}, ],
		fields : {},
		events : {
			depose : "destroy"
		}
	}
};
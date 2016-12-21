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
    }
};
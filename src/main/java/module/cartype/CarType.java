package module.cartype;

import org.nutz.dao.entity.annotation.Comment;
import org.nutz.dao.entity.annotation.Table;

@Table("car_type_new")
@Comment("API表")
public class CarType {
    private String level_id;
    private int ypc_id;
    private String initials;
    private String brand;
    private String factory;
    private String cars;
    private String year;
    private String displacement;
    private String sale_name;
    private String car_type;
    private String year_model;
    private String chassis;
    private String engine;
    private String intake_type;
    private String gearbox_type;
    private String gearbox_remark;
    private String brand_logo_small;

    public String getLevel_id() {
        return level_id;
    }

    public void setLevel_id(String level_id) {
        this.level_id = level_id;
    }

    public int getYpc_id() {
        return ypc_id;
    }

    public void setYpc_id(int ypc_id) {
        this.ypc_id = ypc_id;
    }

    public String getInitials() {
        return initials;
    }

    public void setInitials(String initials) {
        this.initials = initials;
    }

    public String getBrand() {
        return brand;
    }

    public void setBrand(String brand) {
        this.brand = brand;
    }

    public String getFactory() {
        return factory;
    }

    public void setFactory(String factory) {
        this.factory = factory;
    }

    public String getCars() {
        return cars;
    }

    public void setCars(String cars) {
        this.cars = cars;
    }

    public String getYear() {
        return year;
    }

    public void setYear(String year) {
        this.year = year;
    }

    public String getDisplacement() {
        return displacement;
    }

    public void setDisplacement(String displacement) {
        this.displacement = displacement;
    }

    public String getSale_name() {
        return sale_name;
    }

    public void setSale_name(String sale_name) {
        this.sale_name = sale_name;
    }

    public String getCar_type() {
        return car_type;
    }

    public void setCar_type(String car_type) {
        this.car_type = car_type;
    }

    public String getYear_model() {
        return year_model;
    }

    public void setYear_model(String year_model) {
        this.year_model = year_model;
    }

    public String getChassis() {
        return chassis;
    }

    public void setChassis(String chassis) {
        this.chassis = chassis;
    }

    public String getEngine() {
        return engine;
    }

    public void setEngine(String engine) {
        this.engine = engine;
    }

    public String getIntake_type() {
        return intake_type;
    }

    public void setIntake_type(String intake_type) {
        this.intake_type = intake_type;
    }

    public String getGearbox_type() {
        return gearbox_type;
    }

    public void setGearbox_type(String gearbox_type) {
        this.gearbox_type = gearbox_type;
    }

    public String getGearbox_remark() {
        return gearbox_remark;
    }

    public void setGearbox_remark(String gearbox_remark) {
        this.gearbox_remark = gearbox_remark;
    }

    public String getBrand_logo_small() {
        return brand_logo_small;
    }

    public void setBrand_logo_small(String brand_logo_small) {
        this.brand_logo_small = brand_logo_small;
    }

}

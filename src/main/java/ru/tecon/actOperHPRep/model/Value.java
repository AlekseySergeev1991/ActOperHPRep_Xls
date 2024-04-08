package ru.tecon.actOperHPRep.model;

import java.io.Serializable;
import java.util.StringJoiner;

public class Value implements Serializable {

    private String zone;
    private String tnv;
    private String tnvGmc;
    private String min;
    private String max;
    private String parValue;
    private String color;
    private String measure;

    public Value(String parValue, String color) {
        this.parValue = parValue;
        this.color = color;
    }

    public Value(String zone, String tnv, String tnvGmc, String min, String max, String parValue, String color) {
        this.zone = zone;
        this.tnv = tnv;
        this.tnvGmc = tnvGmc;
        this.min = min;
        this.max = max;
        this.parValue = parValue;
        this.color = color;
    }

    public Value(String zone, String min, String max, String parValue, String color) {
        this.zone = zone;
        this.min = min;
        this.max = max;
        this.parValue = parValue;
        this.color = color;
    }

    public Value(String zone, String min, String max, String parValue, String color, String measure) {
        this.zone = zone;
        this.min = min;
        this.max = max;
        this.parValue = parValue;
        this.color = color;
        this.measure = measure;
    }

    public String getZone() {
        return zone;
    }

    public void setZone(String zone) {
        this.zone = zone;
    }

    public String getTnv() {
        return tnv;
    }

    public void setTnv(String tnv) {
        this.tnv = tnv;
    }

    public String getTnvGmc() {
        return tnvGmc;
    }

    public void setTnvGmc(String tnvGmc) {
        this.tnvGmc = tnvGmc;
    }

    public String getMin() {
        return min;
    }

    public void setMin(String min) {
        this.min = min;
    }

    public String getMax() {
        return max;
    }

    public void setMax(String max) {
        this.max = max;
    }

    public String getParValue() {
        return parValue;
    }

    public void setParValue(String parValue) {
        this.parValue = parValue;
    }

    public String getColor() {
        return color;
    }

    public void setColor(String color) {
        this.color = color;
    }

    public String getMeasure() {
        return measure;
    }

    public void setMeasure(String measure) {
        this.measure = measure;
    }

    @Override
    public String toString() {
        return new StringJoiner(", ", Value.class.getSimpleName() + "[", "]")
                .add("zone='" + zone + "'")
                .add("tnv='" + tnv + "'")
                .add("tnvGmc='" + tnvGmc + "'")
                .add("min='" + min + "'")
                .add("max='" + max + "'")
                .add("parValue='" + parValue + "'")
                .add("color='" + color + "'")
                .add("measure='" + measure + "'")
                .toString();
    }
}

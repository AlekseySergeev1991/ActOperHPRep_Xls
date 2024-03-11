package ru.tecon.actOperHPRep.model;

import java.io.Serializable;
import java.util.List;
import java.util.StringJoiner;

public class ReportObject implements Serializable {
    private int numPP;
    private int objId;
    private String objName;
    private String filial;
    private String predpr;
    private String objAddress;
    private List<Value> values;

    public ReportObject(int numPP, int objId, String objName, String filial, String predpr, String objAddress) {
        this.numPP = numPP;
        this.objId = objId;
        this.objName = objName;
        this.filial = filial;
        this.predpr = predpr;
        this.objAddress = objAddress;
    }

    public int getNumPP() {
        return numPP;
    }

    public void setNumPP(int numPP) {
        this.numPP = numPP;
    }

    public int getObjId() {
        return objId;
    }

    public void setObjId(int objId) {
        this.objId = objId;
    }

    public String getObjName() {
        return objName;
    }

    public void setObjName(String objName) {
        this.objName = objName;
    }

    public String getFilial() {
        return filial;
    }

    public void setFilial(String filial) {
        this.filial = filial;
    }

    public String getPredpr() {
        return predpr;
    }

    public void setPredpr(String predpr) {
        this.predpr = predpr;
    }

    public String getObjAddress() {
        return objAddress;
    }

    public void setObjAddress(String objAddress) {
        this.objAddress = objAddress;
    }

    public List<Value> getValues() {
        return values;
    }

    public void setValues(List<Value> values) {
        this.values = values;
    }

    @Override
    public String toString() {
        return new StringJoiner(", ", ReportObject.class.getSimpleName() + "[", "]")
                .add("numPP=" + numPP)
                .add("objId=" + objId)
                .add("objName='" + objName + "'")
                .add("filial='" + filial + "'")
                .add("predpr='" + predpr + "'")
                .add("objAddress='" + objAddress + "'")
                .add("values=" + values)
                .toString();
    }
}

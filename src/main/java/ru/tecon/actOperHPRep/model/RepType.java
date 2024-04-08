package ru.tecon.actOperHPRep.model;

import java.io.Serializable;
import java.time.LocalDateTime;
import java.util.StringJoiner;

public class RepType implements Serializable {

    private LocalDateTime beg;
    private LocalDateTime end;
    private String type;
    private String interval;
    private String typeCode;

    public RepType() {
    }

    public LocalDateTime getBeg() {
        return beg;
    }

    public void setBeg(LocalDateTime beg) {
        this.beg = beg;
    }

    public LocalDateTime getEnd() {
        return end;
    }

    public void setEnd(LocalDateTime end) {
        this.end = end;
    }

    public String getType() {
        return type;
    }

    public void setType(String type) {
        this.type = type;
    }

    public String getInterval() {
        return interval;
    }

    public void setInterval(String interval) {
        this.interval = interval;
    }

    public String getTypeCode() {
        return typeCode;
    }

    public void setTypeCode(String typeCode) {
        this.typeCode = typeCode;
    }

    @Override
    public String toString() {
        return new StringJoiner(", ", RepType.class.getSimpleName() + "[", "]")
                .add("beg=" + beg)
                .add("end=" + end)
                .add("type='" + type + "'")
                .add("interval='" + interval + "'")
                .add("typeCode='" + typeCode + "'")
                .toString();
    }
}

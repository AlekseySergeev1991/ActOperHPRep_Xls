package ru.tecon.actOperHPRep.ejb;

import ru.tecon.actOperHPRep.ActOperHPRep;

import javax.annotation.Resource;
import javax.ejb.Stateless;
import javax.sql.DataSource;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.logging.Logger;

@Stateless
public class ActOperHPRepBean {

    private static final Logger LOGGER = Logger.getLogger(ActOperHPRepBean.class.getName());

    @Resource(name = "jdbc/DataSourceR")
    private DataSource dsR;
    @Resource(name = "jdbc/DataSource")
    private DataSource dsRW;


    public void createReport(int reportId) {
        ExecutorService executor = Executors.newSingleThreadExecutor();
        executor.execute(() -> {
            try {
                ActOperHPRep.makeReport(reportId, dsR, dsRW);
            } catch (InterruptedException e) {
                throw new RuntimeException(e);
            }
        });
    }
}



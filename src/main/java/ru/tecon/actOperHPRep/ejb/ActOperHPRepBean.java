package ru.tecon.actOperHPRep.ejb;

import jakarta.annotation.Resource;
import jakarta.ejb.Stateless;
import ru.tecon.actOperHPRep.ActOperHPRep;

import javax.sql.DataSource;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.logging.Logger;

@Stateless
public class ActOperHPRepBean {

    private static final Logger LOGGER = Logger.getLogger(ActOperHPRepBean.class.getName());

    @Resource(name = "jdbc/DataSource")
    private DataSource dsRw;

    @Resource(name = "jdbc/DataSourceR")
    private DataSource dsR;


    public void createReport(int reportId) {
        ExecutorService executor = Executors.newSingleThreadExecutor();
        executor.execute(() -> {
            try {
                ActOperHPRep.makeReport(reportId, dsR, dsRw);
            } catch (InterruptedException e) {
                throw new RuntimeException(e);
            }
        });
    }
}



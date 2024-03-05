package ru.tecon.actOperHPRep;

import ru.tecon.actOperHPRep.ejb.ActOperHPRepBean;

import javax.ejb.EJB;
import javax.servlet.ServletException;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;

@WebServlet("/loadActOperHPRep")
public class Servlet extends HttpServlet {

    @EJB
    private ActOperHPRepBean actOperHPRepBean;

    @Override
    protected void doGet(HttpServletRequest req, HttpServletResponse resp) throws ServletException, IOException {
        int reportId = Integer.parseInt(req.getParameter("reportId"));

        actOperHPRepBean.createReport(reportId);


        resp.setStatus(HttpServletResponse.SC_OK);
    }

}

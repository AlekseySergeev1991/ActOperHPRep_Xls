package ru.tecon.actOperHPRep;

import jakarta.ejb.EJB;
import jakarta.servlet.ServletException;
import jakarta.servlet.annotation.WebServlet;
import jakarta.servlet.http.HttpServlet;
import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpServletResponse;
import ru.tecon.actOperHPRep.ejb.ActOperHPRepBean;

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

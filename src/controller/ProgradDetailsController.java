package controller;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import javax.servlet.RequestDispatcher;
import javax.servlet.ServletException;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import model.Prograd;
import service.ExcelGenerator;

@WebServlet(urlPatterns = { "/prograd" })
public class ProgradDetailsController extends HttpServlet {
	private static final long serialVersionUID = 1L;
// Uncomment the below line
	List<Prograd> list = new ArrayList<Prograd>();
	RequestDispatcher rd;
	public ProgradDetailsController() {
		super();
	}

	protected void doGet(HttpServletRequest request, HttpServletResponse response)
			throws ServletException, IOException {
		System.out.println("Entering do get");
		rd = this.getServletContext().getRequestDispatcher("/WEB-INF/views/index.jsp");
		rd.forward(request, response);
	}

	protected void doPost(HttpServletRequest request, HttpServletResponse response)
			throws ServletException, IOException {
		// uncomment the code below to see the final output
		String name = request.getParameter("name");
		String id = request.getParameter("id");
		String rate=request.getParameter("rating");
		String recom1 = request.getParameter("re1");
		String recom2 = request.getParameter("re2");
		String recom3 = request.getParameter("re3");
		String recom4 = request.getParameter("re4");
		String recom5 = request.getParameter("re5");
		String comment = request.getParameter("comment");
		String submit = request.getParameter("submit");
		String download = request.getParameter("download");
		
		
		Prograd prograd = new Prograd(name, id, rate, recom1, comment);


		if (recom1 != null)
			prograd.setRecommend(recom1);
		if (recom2 != null)
			prograd.setRecommend(recom2);
		if (recom3 != null)
			prograd.setRecommend(recom3);
		if (recom4 != null)
			prograd.setRecommend(recom4);
		if (recom5 != null)
			prograd.setRecommend(recom5);

		
		prograd.setName(name);
		prograd.setId(id);
		prograd.setComment(comment);
		prograd.setRate(rate);
	
		list.add(prograd);

		if (submit != null) {
			rd = this.getServletContext().getRequestDispatcher("/WEB-INF/views/result.jsp");
			rd.forward(request, response);
		}

		if (download != null) {
			response.setContentType("application/octet-stream");
			response.setHeader("Content-Disposition", "attachment;filename=prograd.xlsx");
		}

		ExcelGenerator excel=new ExcelGenerator();
		HSSFWorkbook hwb = excel.excelGenerate(prograd,list);
		System.out.println(hwb);
			if(hwb!=null)
				request.setAttribute("upload_message", "File downloaded successfully");
				rd = this.getServletContext().getRequestDispatcher("/WEB-INF/views/result.jsp");
				rd.forward(request, response);

	}
}

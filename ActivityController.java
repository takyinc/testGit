package com.bjpowernode.crm.workbench.web.controller;

import com.bjpowernode.crm.commons.constants.Constants;
import com.bjpowernode.crm.commons.domain.ReturnObject;
import com.bjpowernode.crm.commons.utils.DateUtils;
import com.bjpowernode.crm.commons.utils.HSSFUtils;
import com.bjpowernode.crm.commons.utils.UUIDUtils;
import com.bjpowernode.crm.settings.domain.User;
import com.bjpowernode.crm.settings.service.UserService;
import com.bjpowernode.crm.workbench.domain.Activity;
import com.bjpowernode.crm.workbench.domain.ActivityRemark;
import com.bjpowernode.crm.workbench.service.ActivityRemarkService;
import com.bjpowernode.crm.workbench.service.ActivityService;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFCellUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.HttpSession;
import java.io.*;
import java.util.*;

@Controller
public class ActivityController {
    @Autowired
    private UserService userService;

    @Autowired
    private ActivityService activityService;

    @Autowired
    private ActivityRemarkService activityRemarkService;

    @RequestMapping("/workbench/activity/index.do")
    public String index(HttpServletRequest request){
        List<User> userList = userService.queryAllUser();
        request.setAttribute("userList", userList);
        return "workbench/activity/index";
    }

    @RequestMapping("/workbench/activity/savecreateactivity.do")
    //目前看只有insert操作才会直接传Activity对象作为参数
    public @ResponseBody Object saveCreateActivity(Activity activity, HttpSession session){
        //封装参数,前端界面保存的参数里面,对应Activity少了三个数据,所以得先把这三个数据填上
        activity.setId(UUIDUtils.getUUID());
        activity.setCreateTime(DateUtils.formateDateTime(new Date()));
        User user = (User)session.getAttribute(Constants.SESSION_USER);
        activity.setCreateBy(user.getId());

        //创建ReturnObject,因为trycatch中要赋值进去
        ReturnObject returnObject = new ReturnObject();
        //opt + cmd + t快速try catch,一般除了select外都得用trycatch
        try {
            int ret = activityService.saveCreateActivity(activity);
            if (ret > 0) {
                returnObject.setCode(Constants.RETURN_OBJECT_CODE_SUCCESS);
            }else {
                returnObject.setCode(Constants.RETURN_OBJECT_CODE_FAIL);
                returnObject.setMessage("系统繁忙...请稍后再试");
            }
        } catch (Exception e) {
            e.printStackTrace();
            returnObject.setCode(Constants.RETURN_OBJECT_CODE_FAIL);
            returnObject.setMessage("系统繁忙...请稍后再试");
        }
        return returnObject;
    }

    @RequestMapping("/workbench/activity/queryActivityByConditionForPage.do")
    public @ResponseBody Object queryActivityByConditionForPage(String name, String owner, String startDate, String endDate, int pageNo, int pageSize){
        Map<String, Object> map = new HashMap<>();
        map.put("name", name);
        map.put("owner", owner);
        map.put("startDate", startDate);
        map.put("endDate", endDate);
        map.put("beginNo", (pageNo-1) * pageSize);
        map.put("pageSize", pageSize);
        //Controller一般不会调用两次service,事务在service层,因为得保证事务的一致性,不会一成功一失败,但是是select就无所谓了,
        // 如果是insert之类的操作,只能是service层调用多次mapper,不能Controler调用多次service
        List<Activity> activityList = activityService.queryActivityByConditionForPage(map);
//        for (Activity acticity: activityList) {
//            System.out.println(acticity);
//        }
        int totalRows = activityService.queryCountOfActivityByCondition(map);
//        System.out.println(totalRows);
        Map<String, Object> retMap = new HashMap<>();
        retMap.put("activityList", activityList);
        retMap.put("totalRows", totalRows);
        return retMap;
    }

    @RequestMapping("/workbench/activity/deleteActivityByIds.do")
    //Mapper和Service层都是传入ids,Controller没有保持一致是因为需和前端保持一致
    public @ResponseBody Object deleteActivityByIds(String[] id){
        ReturnObject returnObject = new ReturnObject();
        //除select都要用trycatch
        try {
            int ret = activityService.deleteActivityByIds(id);
            if (ret > 0){
                returnObject.setCode(Constants.RETURN_OBJECT_CODE_SUCCESS);
            }else {
                returnObject.setCode(Constants.RETURN_OBJECT_CODE_FAIL);
                returnObject.setMessage("系统繁忙,请稍后再试...");
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return returnObject;
    }

    @RequestMapping("/workbench/activity/queryActivityById.do")
    public @ResponseBody Object queryActivityById(String id){
        Activity activity = activityService.queryActivityById(id);
        return activity;
    }

    @RequestMapping("/workbench/activity/saveEditActivity.do")
    public @ResponseBody Object saveEditActivity(Activity activity, HttpSession session){
        //前端传入Controller后会缺少两个参数,所以先补上参数
        activity.setEditTime(DateUtils.formateDateTime(new Date()));
        User user = (User) session.getAttribute(Constants.SESSION_USER);
        activity.setEditBy(user.getId());
        System.out.println(activity.getEditBy());
        ReturnObject returnObject = new ReturnObject();
        try {
            int ret = activityService.saveEditActivity(activity);
            if (ret > 0) {
                returnObject.setCode(Constants.RETURN_OBJECT_CODE_SUCCESS);
            }else {
                returnObject.setCode(Constants.RETURN_OBJECT_CODE_FAIL);
                returnObject.setMessage("系统繁忙,请稍后再试...");
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return returnObject;
    }

    //这是导出excel test的
    @RequestMapping("/workbench/activity/fileDownload.do")
    //返回值只能是返回实体类或者字符串等信息,文件得用response去返给浏览器
    public void fileDownload(HttpServletResponse response) throws Exception{
        System.out.println("进入controller成功.");
       //设置响应类型,google搜ContentType就有对应表了, excel为application/octet-stream
        response.setContentType("application/octet-stream;charset=UTF-8");
        //PrintWriter为字符流限制大,excel为字节二进制所以不能用
//        PrintWriter out = response.getWriter();
        //获取输出流
        OutputStream out = response.getOutputStream();

        response.addHeader("Content-Disposition", "attachment;filename=mystudentList.xls");

        //读取excel数据(InputStream),然后输出到浏览器(OutputStream)
        InputStream is = new FileInputStream("/Users/takyin/IdeaProjects/Student.xls");
        //字节块读取,快一点
        byte[] buff = new byte[256];
        //这里未理解
        int len = 0;
        while ((len = is.read(buff)) != -1){
            out.write(buff, 0, len);
        }

        //关闭资源,凡是我们自己new的才需要关,这个的OutputStream不需要关,tomcat会自己关,我们手动关有可能报错,因为tomcat其他地方可能还需要用到
        is.close();
        out.flush();
    }

    @RequestMapping("/workbench/activity/exportExcel.do")
    public void exportExcel(HttpServletResponse response) throws Exception{
        //生成excel
        HSSFWorkbook workbook = new HSSFWorkbook();
        //生成表
        HSSFSheet sheet = workbook.createSheet("市场活动列表");
        //生成第一行
        HSSFRow row = sheet.createRow(0);
        //生成具体单元格及赋value
        HSSFCell cell = row.createCell(0);
        cell.setCellValue("ID");
        cell = row.createCell(1);
        cell.setCellValue("所有者");
        cell = row.createCell(2);
        cell.setCellValue("名字");
        cell = row.createCell(3);
        cell.setCellValue("开始日期");
        cell = row.createCell(4);
        cell.setCellValue("结束日期");
        cell = row.createCell(5);
        cell.setCellValue("花费");
        cell = row.createCell(6);
        cell.setCellValue("描述");
        cell = row.createCell(7);
        cell.setCellValue("创建时间");
        cell = row.createCell(8);
        cell.setCellValue("创建人");
        cell = row.createCell(9);
        cell.setCellValue("最后编辑时间");
        cell = row.createCell(10);
        cell.setCellValue("最后编辑人");
        List<Activity> activityList = activityService.queryAllActivity();
        if (activityList != null && activityList.size() != 0) {
            //如果activityList数据不为空,就把数据赋值给对应单元格
            for (int i = 0; i < activityList.size(); i++) {
                Activity activity = activityList.get(i);
                row = sheet.createRow(i + 1);
                cell = row.createCell(0);
                cell.setCellValue(activity.getId());
                cell = row.createCell(1);
                cell.setCellValue(activity.getOwner());
                cell = row.createCell(2);
                cell.setCellValue(activity.getName());
                cell = row.createCell(3);
                cell.setCellValue(activity.getStartDate());
                cell = row.createCell(4);
                cell.setCellValue(activity.getEndDate());
                cell = row.createCell(5);
                cell.setCellValue(activity.getCost());
                cell = row.createCell(6);
                cell.setCellValue(activity.getCost());
                cell = row.createCell(7);
                cell.setCellValue(activity.getCreateTime());
                cell = row.createCell(8);
                cell.setCellValue(activity.getCreateBy());
                cell = row.createCell(9);
                cell.setCellValue(activity.getEditTime());
                cell = row.createCell(10);
                cell.setCellValue(activity.getEditBy());
            }
        }
/*        此种代码为先内存->磁盘写出文件后,再磁盘->内存这样的方式给客户端,凡是动用到磁盘效率都会很低,java提供了直接内存到内存的方法
          这个会在本机指定的位置先生成一个文件(输入流),然后再用输入流拿到这个文件,再用response输出流给浏览器
        //根据wb对象生成excel文件,生成excel的路径存放位置,和浏览器下载到哪里是无关的
        OutputStream os = new FileOutputStream("/Users/takyin/Desktop/acticityList.xls");
        workbook.write(os);
        //关闭资源
        os.close();
        workbook.close();

        //把生成的excel文件下载到客户端
        response.setContentType("application/octet-stream;charset=UTF-8");
        response.addHeader("Content-Disposition", "attachment;filename=acticityList.xls");
        InputStream is = new FileInputStream("/Users/takyin/Desktop/acticityList.xls");
        OutputStream out = response.getOutputStream();
        byte[] buff = new byte[256];
        int len = 0;
        while ((len = is.read(buff)) != -1){
            out.write(buff, 0, len);
        }
        is.close();
        out.flush();*/

        //这种方式就不会在磁盘生成一个文件,直接内存到内存
        //指定response的类型为excel
        response.setContentType("application/octet-stream;charset=UTF-8");
        //增加一个Header,让浏览器不要自动打开文件,而是调用下载框,并对文件赋予默认名字
        response.addHeader("Content-Disposition", "attachment;filename=activityList.xls");
        //利用response拿到给浏览器输出的输出流
        OutputStream out = response.getOutputStream();
        //workbook写入一个输出流,这个是out传入给write,有点不太好理解
        workbook.write(out);
        workbook.close();
        out.flush();
    }

    @RequestMapping("/workbench/activity/exportExcelByIds.do")
    public void exportExcelByIds(String[] id, HttpServletResponse response) throws Exception{
        //生成excel
        HSSFWorkbook workbook = new HSSFWorkbook();
        //生成表
        HSSFSheet sheet = workbook.createSheet("市场活动列表");
        //生成第一行
        HSSFRow row = sheet.createRow(0);
        //生成具体单元格及赋value
        HSSFCell cell = row.createCell(0);
        cell.setCellValue("ID");
        cell = row.createCell(1);
        cell.setCellValue("所有者");
        cell = row.createCell(2);
        cell.setCellValue("名字");
        cell = row.createCell(3);
        cell.setCellValue("开始日期");
        cell = row.createCell(4);
        cell.setCellValue("结束日期");
        cell = row.createCell(5);
        cell.setCellValue("花费");
        cell = row.createCell(6);
        cell.setCellValue("描述");
        cell = row.createCell(7);
        cell.setCellValue("创建时间");
        cell = row.createCell(8);
        cell.setCellValue("创建人");
        cell = row.createCell(9);
        cell.setCellValue("最后编辑时间");
        cell = row.createCell(10);
        cell.setCellValue("最后编辑人");
        List<Activity> activityList = activityService.queryActivityByIds(id);
        if (activityList != null && activityList.size() != 0) {
            //如果activityList数据不为空,就把数据赋值给对应单元格
            for (int i = 0; i < activityList.size(); i++) {
                Activity activity = activityList.get(i);
                row = sheet.createRow(i + 1);
                cell = row.createCell(0);
                cell.setCellValue(activity.getId());
                cell = row.createCell(1);
                cell.setCellValue(activity.getOwner());
                cell = row.createCell(2);
                cell.setCellValue(activity.getName());
                cell = row.createCell(3);
                cell.setCellValue(activity.getStartDate());
                cell = row.createCell(4);
                cell.setCellValue(activity.getEndDate());
                cell = row.createCell(5);
                cell.setCellValue(activity.getCost());
                cell = row.createCell(6);
                cell.setCellValue(activity.getCost());
                cell = row.createCell(7);
                cell.setCellValue(activity.getCreateTime());
                cell = row.createCell(8);
                cell.setCellValue(activity.getCreateBy());
                cell = row.createCell(9);
                cell.setCellValue(activity.getEditTime());
                cell = row.createCell(10);
                cell.setCellValue(activity.getEditBy());
            }
        }
        //这种方式就不会在磁盘生成一个文件,直接内存到内存
        //指定response的类型为excel
        response.setContentType("application/octet-stream;charset=UTF-8");
        //增加一个Header,让浏览器不要自动打开文件,而是调用下载框,并对文件赋予默认名字
        response.addHeader("Content-Disposition", "attachment;filename=activityList.xls");
        //利用response拿到给浏览器输出的输出流
        OutputStream out = response.getOutputStream();
        //workbook写入一个输出流,这个是out传入给write,有点不太好理解
        workbook.write(out);
        workbook.close();
        out.flush();
    }

    @RequestMapping("/workbench/activity/fileUpload.do")
    public @ResponseBody Object fileUpload(MultipartFile activityFile, String userName, HttpSession session) throws Exception{
        System.out.println("userName=" + userName);
        File file = new File("/Users/takyin/Desktop/aaa.xls");
        activityFile.transferTo(file);

        ReturnObject returnObject = new ReturnObject();
        returnObject.setCode(Constants.RETURN_OBJECT_CODE_SUCCESS);
        returnObject.setMessage("上传成功");
        return returnObject;
    }
    @RequestMapping("/workbench/activity/importByExcel.do")
    public @ResponseBody Object importByExcel(MultipartFile activityFile, String userName, HttpSession session){
        System.out.println("进入controller成功");
        System.out.println("userName=" + userName);
        User user = (User) session.getAttribute(Constants.SESSION_USER);
        ReturnObject returnObject = new ReturnObject();
        try {
            //将前端拿到的excel生成到本机(服务器)的指定位置,这种也是内存---磁盘---内存,所以得优化
//            String originalFilename = activityFile.getOriginalFilename();
//            File file = new File("/Users/takyin/IdeaProjects" + originalFilename);
//            activityFile.transferTo(file);
            //优化后,直接用activityFile拿到输入流放入HSSFWorkbook即可
            InputStream is = activityFile.getInputStream();
            //解析拿到的excel文件中的数据
//            InputStream is = new FileInputStream("/Users/takyin/IdeaProjects" + originalFilename);
            HSSFWorkbook wb = new HSSFWorkbook(is);
            HSSFSheet sheet = wb.getSheetAt(0);
            HSSFRow row = null;
            HSSFCell cell = null;
            Activity activity = null;
            List<Activity> activityList = new ArrayList<>();
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                row = sheet.getRow(i);
                activity = new Activity();
                activity.setId(UUIDUtils.getUUID());
                activity.setOwner(user.getId());
                activity.setCreateTime(DateUtils.formateDateTime(new Date()));
                activity.setCreateBy(user.getId());
                for (int j = 0; j < row.getLastCellNum(); j++) {
                    cell = row.getCell(j);
                    if (j == 0) {
                        activity.setName(HSSFUtils.getCellValueForStr(cell));
                    }else if(j == 1){
                        activity.setStartDate(HSSFUtils.getCellValueForStr(cell));
                    }else if(j == 2){
                        activity.setEndDate(HSSFUtils.getCellValueForStr(cell));
                    }else if(j == 3){
                        activity.setCost(HSSFUtils.getCellValueForStr(cell));
                    }else if(j == 4){
                        activity.setDescription(HSSFUtils.getCellValueForStr(cell));
                    }
                }
                activityList.add(activity);
            }
            int ret = activityService.saveActivityByList(activityList);
            //这里没有用到if来判断ret是否>0,因为ret为0时,可能只是excel无数据,不属于fail,所有只有exception才证明失败
            returnObject.setCode(Constants.RETURN_OBJECT_CODE_SUCCESS);
            returnObject.setRetData(ret);
        } catch (Exception e) {
            e.printStackTrace();
            returnObject.setCode(Constants.RETURN_OBJECT_CODE_FAIL);
            returnObject.setMessage("系统繁忙,请稍后再试试...");
        }
        return returnObject;
    }

//    @RequestMapping("/workbench/activity/detailActivity.do")
//    public String detailActivity(String id, HttpServletRequest request){
//        Activity activity = activityService.queryActivityForDetailById(id);
//        List<ActivityRemark> activityRemarkList = activityRemarkService.queryActivityRemarkForDetailByActivityId(id);
//        request.setAttribute("activity", activity);
//        request.setAttribute("activityRemarkList", activityRemarkList);
//        return "workbench/activity/detail";
//    }
}

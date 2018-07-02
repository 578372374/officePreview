package com.document.listener;

import javax.servlet.ServletContextEvent;
import javax.servlet.ServletContextListener;

import com.document.bean.Constant;

/**
 * 监听tomcat服务器启动/停止事件
 *
 */
public class InitServletContextListener implements ServletContextListener {

	public void contextInitialized(ServletContextEvent sce) {
		//设置web应用名称
		Constant.APPLICATION_NAME = sce.getServletContext().getContextPath();
		//设置web应用/的绝对路径
		Constant.REAL_PATH = sce.getServletContext().getRealPath("/");
	}

	public void contextDestroyed(ServletContextEvent sce) {

	}
}

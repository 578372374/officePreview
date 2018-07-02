package com.document.listener;

import javax.servlet.ServletContextEvent;
import javax.servlet.ServletContextListener;

import com.document.bean.JodConverterContext;

/**
 * 
* 用于初始化JodConverter
 */
public class JodConverterListener implements ServletContextListener{
	@Override
	public void contextInitialized(ServletContextEvent envent) {
		JodConverterContext.init(envent.getServletContext());
	}

	//如果想容器关闭的时候，执行这个方法，需要调用shutdown.bat来关闭tomcat才可以
	@Override
	public void contextDestroyed(ServletContextEvent arg0) {
		JodConverterContext.destroy();
	}
}

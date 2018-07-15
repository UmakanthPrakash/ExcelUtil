package com.jaxrxapachepoi.rest.config;

import javax.ws.rs.ApplicationPath;

import org.glassfish.jersey.server.ResourceConfig;
import org.springframework.stereotype.Component;

@Component
@ApplicationPath("/")
public class SpringDemoConfig extends ResourceConfig{
	public SpringDemoConfig(){
		packages("com.springdemo.rest.controller");
	}
}

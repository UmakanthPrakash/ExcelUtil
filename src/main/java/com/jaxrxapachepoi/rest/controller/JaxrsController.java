package com.jaxrxapachepoi.rest.controller;

import javax.ws.rs.GET;
import javax.ws.rs.Path;
import javax.ws.rs.PathParam;

import org.springframework.stereotype.Component;

@Component
@Path("/greetings")
public class JaxrsController {

	@GET
	@Path("{name}")
	public String getUserName(@PathParam("id") String name){
		return "Greetings " + name;
	}
}

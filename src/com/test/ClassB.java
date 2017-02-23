package com.test;

public class ClassB {
	int a=0;
	
	public void disp()
	{
		a=1;
	}
	
	public synchronized void disp1()
	{
		System.out.println(a);
	}
	
	
}

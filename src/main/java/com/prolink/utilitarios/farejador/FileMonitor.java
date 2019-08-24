package com.prolink.utilitarios.farejador;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.nio.file.attribute.FileOwnerAttributeView;
import java.util.*;

import net.contentobjects.jnotify.JNotify;
import net.contentobjects.jnotify.JNotifyListener;

public class FileMonitor {
	public static void main(String[] args) throws Exception{
		// path to watch
	    List<String> array = Arrays.asList(new String[] {
	    	"\\\\plkserver\\Todos Departamentos",
		    "\\\\plkserver\\Obrigacoes",
		    "\\\\plkserver\\Regularizacao",
		    "\\\\plkserver\\Fiscal",
		    "\\\\plkserver\\Contabil"
	    });
	    // watch mask, specify events you care about,
	    // or JNotify.FILE_ANY for all events.
	    int mask = JNotify.FILE_CREATED  | 
	               JNotify.FILE_DELETED  | 
	               JNotify.FILE_MODIFIED | 
	               JNotify.FILE_RENAMED;
	
	    // watch subtree?
	    boolean watchSubtree = true;
	    
	    // add actual watch
	    for(String s :array) {
		    int watchID1 = JNotify.addWatch(s, mask, watchSubtree, new NewListener());
		    // sleep a little, the application will exit if you
		    // don't (watching is asynchronous), depending on your
		    // application, this may not be required
		    System.out.println(watchID1);
	    }
	    Thread.sleep(10000000);
	    // to remove watch the watch
	    //if (!JNotify.removeWatch(watchID)) {
	      // invalid watch ID specified.
	    //}
  }
  public static class NewListener implements JNotifyListener {
	public void fileRenamed(int wd, String rootPath, String oldName,
        String newName) {
	    FileOwnerAttributeView foav = Files.getFileAttributeView(Paths.get(rootPath+"\\"+newName),FileOwnerAttributeView.class);
	      try {
			print("renamed " + rootPath + " : " + oldName + " -> " + newName+" by "+foav.getOwner().getName());
		} catch (IOException e) {
			e.printStackTrace();
		}
    }
    public void fileModified(int wd, String rootPath, String name) {
    	FileOwnerAttributeView foav = Files.getFileAttributeView(Paths.get(rootPath+"\\"+name),FileOwnerAttributeView.class);
	    try {
	    		print("modified " + rootPath + " : " + name+" by "+foav.getOwner().getName());
	        } catch (IOException e) {
		    	e.printStackTrace();
		}
    }
    public void fileDeleted(int wd, String rootPath, String name) {
      print("deleted " + rootPath + " : " + name);
    }
    public void fileCreated(int wd, String rootPath, String name) {
      	FileOwnerAttributeView foav = Files.getFileAttributeView(Paths.get(rootPath+"\\"+name),FileOwnerAttributeView.class);
	    try {
	    	print("created " + rootPath + " : " + name+" by "+foav.getOwner().getName());
	    } catch (IOException e) {
		    	e.printStackTrace();
	    }
    }
    void print(String msg) {
      System.err.println(msg);
    }
  }
}

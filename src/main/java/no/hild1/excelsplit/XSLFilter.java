package no.hild1.excelsplit;

import java.io.File;

import javax.swing.filechooser.FileFilter;

public class XSLFilter extends FileFilter {

	public static String getExtension(File f) {
		String ext = null;
		String s = f.getName();
		int i = s.lastIndexOf('.');

		if (i > 0 &&  i < s.length() - 1) {
			ext = s.substring(i+1).toLowerCase();
		}
		return ext;
	}

	//Accept all directories and all xls files
	public boolean accept(File f) {
		if (f.isDirectory()) {
			return true;
		}

		String extension = getExtension(f);
		if (extension != null) {
			if (extension.equals("xls")) {
				return true;
			} else {
				return false;
			}
		}
		return false;
	}

	//The description of this filter
	public String getDescription() {
		return "Microsoft XSL-regneark";
	}
}


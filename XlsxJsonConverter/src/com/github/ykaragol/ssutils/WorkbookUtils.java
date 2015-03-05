package com.github.ykaragol.ssutils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.security.GeneralSecurityException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class WorkbookUtils {

	public Workbook createWorkbook(String fileName) {
		File file = new File(fileName);
		return createWorkbook(file);
	}

	public Workbook createWorkbook(File file) {

		POIFSFileSystem filesystem;
		try {
			filesystem = new POIFSFileSystem(new FileInputStream(file));
			EncryptionInfo info = new EncryptionInfo(filesystem);
			Decryptor d = Decryptor.getInstance(info);

			if (!d.verifyPassword(Decryptor.DEFAULT_PASSWORD)) {
				throw new RuntimeException(
						"Unable to process: document is encrypted");
			}

			InputStream dataStream = d.getDataStream(filesystem);

			return WorkbookFactory.create(dataStream);

		} catch (GeneralSecurityException ex) {
			throw new RuntimeException("Unable to process encrypted document",
					ex);
		} catch (InvalidFormatException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return null;

		// Workbook wb = null;
		// try {
		// wb = WorkbookFactory.create(file);
		// } catch (InvalidFormatException e) {
		// e.printStackTrace();
		// } catch (IOException e) {
		// e.printStackTrace();
		// }
		// return wb;
	}

}

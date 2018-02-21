/**
 * 
 */
package com.funny.office.service;

import com.funny.office.po.ScormMeta;
import org.apache.commons.io.FileUtils;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;
import org.w3c.dom.Document;
import org.w3c.dom.Node;
import org.xml.sax.SAXException;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathConstants;
import javax.xml.xpath.XPathExpressionException;
import javax.xml.xpath.XPathFactory;
import java.io.*;
import java.nio.charset.Charset;
import java.util.Enumeration;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;

/**
 * @author funnyZpC Scorm文件处理
 */
@Service
@Transactional(rollbackFor=Throwable.class)
public class ScormService {

	/**
	 * 获取SCOMR的元数据
	 * 
	 * @param path
	 *            课件包的地址，可以是一个zip文件，也可以是一个解开的课件目录
	 * @return
	 * @throws IOException
	 * @throws ParserConfigurationException
	 * @throws SAXException
	 * @throws XPathExpressionException
	 */
	public ScormMeta getMeta(String path)
			throws IOException, SAXException, ParserConfigurationException, XPathExpressionException {
		ScormMeta meta = new ScormMeta();
		InputStream is = getManifest(path);
		Document doc = DocumentBuilderFactory.newInstance().newDocumentBuilder().parse(is);
		XPath xp = XPathFactory.newInstance().newXPath();
		Node no = (Node) xp.evaluate("/manifest/organizations", doc, XPathConstants.NODE);
		Node defaultOrgnization = no.getAttributes().getNamedItem("default");
		if (defaultOrgnization == null) {
			no = (Node) xp.evaluate("organization", no, XPathConstants.NODE);
		} else {
			String tmp = defaultOrgnization.getNodeValue();
			no = (Node) xp.evaluate("organization[@identifier=\"" + tmp + "\"]", no, XPathConstants.NODE);
		}

		Node ni = (Node) xp.evaluate("item[@identifier=\"ITEM\"]", no, XPathConstants.NODE);
		if (ni == null)
			ni = no;
		meta.setTitle(xp.evaluate("title", ni));
		ni = (Node) xp.evaluate("item", ni, XPathConstants.NODE);
		String id = ni.getAttributes().getNamedItem("identifierref").getNodeValue();

		Node nr = (Node) xp.evaluate("//resources/resource[@identifier=\"" + id + "\"]", doc, XPathConstants.NODE);
		meta.setHref(nr.getAttributes().getNamedItem("href").getNodeValue());		
		is.close();
		return meta;
	}

	private InputStream getManifest(String path) throws IOException {
		File f = new File(path);
		if (!f.exists())
			throw new IOException(String.format("[%s]不存在", path));

		// fold
		if (f.isDirectory()) {
			f = new File(f.getAbsolutePath() + File.separator + "imsmanifest.xml");
			if (!f.exists())
				throw new IOException(String.format("[%s]不存在", f.getAbsolutePath()));
			return new FileInputStream(f);
		} else {
			ZipFile zf= new ZipFile(f);
			try{
				ZipEntry entry = zf.getEntry("imsmanifest.xml");
				ByteArrayOutputStream baos = new ByteArrayOutputStream();
				InputStream is =zf.getInputStream(entry);
				int buf = 0;
				while((buf = is.read()) != -1)
					baos.write(buf);
				ByteArrayInputStream bais = new ByteArrayInputStream(baos.toByteArray());
				baos.close();
				return bais;
			}finally{
				zf.close();
			}
		}
	}

	/**
	 * 将目录压缩到zip文件
	 * 
	 * @param fold
	 * @param file
	 * @throws IOException
	 */
	public void zip(String fold, String file) throws IOException {
		File fSrc = new File(fold);
		File fDest = new File(file);
		if (!fSrc.exists() || !fSrc.isDirectory())
			throw new IOException(String.format("Source [%s] is not exist or not a directory.", fold));

		if (fDest.exists())
			fDest.delete();
		
		if (!fDest.getParentFile().exists())
			FileUtils.forceMkdir(fDest.getParentFile());

		ZipOutputStream zos = new ZipOutputStream(new FileOutputStream(file));
		try {
			zip(zos, fSrc, "");
		} finally {
			zos.close();
		}
	}

	private void zip(ZipOutputStream zos, File fSrc, String base) throws IOException {
		if (fSrc.isDirectory()) {
			File[] subFile = fSrc.listFiles();
			zos.putNextEntry(new ZipEntry(base + "/"));
			base = (base.length() == 0) ? "" : (base + "/");

			for (int i = 0; i < subFile.length; i++) {
				zip(zos, subFile[i], base + subFile[i].getName());
			}
		} else {
			zos.putNextEntry(new ZipEntry(base));
			BufferedInputStream bufferInputStream = new BufferedInputStream(new FileInputStream(fSrc));
			try {
				int buf;
				while ((buf = bufferInputStream.read()) != -1) {
					zos.write(buf);
				}
			} finally {
				bufferInputStream.close();
			}
		}
	}

	/**
	 * 解压缩zip文件
	 * @param file
	 * @param fold
	 * @throws IOException
	 */
	public void unzip(String file, String fold) throws IOException {
		File fSrc = new File(file);
		File fDest = new File(fold);
		if (!fSrc.exists())
			throw new IOException(String.format("Source [%s] is not exist.", file));
		if (fDest.exists())
			FileUtils.forceDelete(fDest);

		Charset gbk = Charset.forName("GBK");
		ZipFile zipFile = new ZipFile(fSrc,gbk);//设置zip包的编码为gbk
		Enumeration emu = zipFile.entries();
		while (emu.hasMoreElements()) {
			ZipEntry entry = (ZipEntry) emu.nextElement();
			// 会把目录作为一个file读出一次，所以只建立目录就可以，之下的文件还会被迭代到。
			if (entry.isDirectory()) {
				new File(fold + entry.getName()).mkdirs();
				continue;
			}
			BufferedInputStream bis = new BufferedInputStream(zipFile.getInputStream(entry));
			File destFile = new File(fold + entry.getName());
			// 加入这个的原因是zipfile读取文件是随机读取的，这就造成可能先读取一个文件
			// 而这个文件所在的目录还没有出现过，所以要建出目录来。
			File parent = destFile.getParentFile();
			if (parent != null && (!parent.exists())) {
				FileUtils.forceMkdir(parent);
			}
			BufferedOutputStream bos = new BufferedOutputStream(new FileOutputStream(destFile));
			try {
				int buf = 0;
				while ((buf = bis.read()) != -1) {
					bos.write(buf);
				}
				bos.flush();
			} finally {
				bos.close();
			}
		}
		zipFile.close();
	}
}

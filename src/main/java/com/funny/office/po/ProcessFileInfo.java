package com.funny.office.po;

public class ProcessFileInfo {

    private Boolean status;//true:成功 false:失败

    private String pkgName;//转换后包的全名

    private String pkgPath;//转换后包的路径

    /**
     *
     * @param status		返回状态 True:成功 False:失败
     * @param pkgName	ZIP包的名称
     * @param pkgPath	ZIP包的路径
     */
    public ProcessFileInfo(Boolean status, String pkgName, String pkgPath) {
        super();
        this.status = status;
        this.pkgName = pkgName;
        this.pkgPath = pkgPath;
    }

    /**
     * @return the status
     */
    public Boolean getStatus() {
        return status;
    }

    /**
     * @param status the status to set
     */
    public void setStatus(Boolean status) {
        this.status = status;
    }

    /**
     * @return the pkgName
     */
    public String getPkgName() {
        return pkgName;
    }

    /**
     * @param pkgName the pkgName to set
     */
    public void setPkgName(String pkgName) {
        this.pkgName = pkgName;
    }

    /**
     * @return the pkgPath
     */
    public String getPkgPath() {
        return pkgPath;
    }

    /**
     * @param pkgPath the pkgPath to set
     */
    public void setPkgPath(String pkgPath) {
        this.pkgPath = pkgPath;
    }


}

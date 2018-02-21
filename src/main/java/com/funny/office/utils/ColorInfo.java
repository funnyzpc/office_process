package com.funny.office.utils;

public class ColorInfo{
    /**
     * ��ɫ��alphaֵ����ֵ��������ɫ��͸����
     */
    public int A;
    /**
     * ��ɫ�ĺ����ֵ��Red
     */
    public int R;
    /**
     * ��ɫ���̷���ֵ��Green
     */
    public int G;
    /**
     * ��ɫ��������ֵ��Blue
     */
    public int B;

    public int toRGB() {
        return this.R << 16 | this.G << 8 | this.B;
    }

    public java.awt.Color toAWTColor(){
        return new java.awt.Color(this.R,this.G,this.B,this.A);
    }

    public static ColorInfo fromARGB(int red, int green, int blue) {
        return new ColorInfo((int) 0xff, (int) red, (int) green, (int) blue);
    }
    public static ColorInfo fromARGB(int alpha, int red, int green, int blue) {
        return new ColorInfo(alpha, red, green, blue);
    }
    public ColorInfo(int a,int r, int g , int b ) {
        this.A = a;
        this.B = b;
        this.R = r;
        this.G = g;
    }
}
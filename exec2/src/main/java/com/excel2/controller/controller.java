package com.excel2.controller;

import java.math.BigDecimal;
import java.text.DecimalFormat;

public class controller {
    private static final String[] tensNames = { "", " mười", " hai mươi", " ba mươi", " bốn mươi", " năm mươi",
            " sáu mươi", " bảy mươi", " tám mươi", " chín mươi" };

    private static final String[] numNames = { "", " một", " hai", " ba", " bốn", " năm", " sáu", " bảy", " tám",
            " chín", " mười", " mười một", " mười hai", " mười ba", " mười bốn", " mười lăm", " mười sáu", " mười bảy",
            " mười tám", " mười chín" };


    public static String convertLessThanOneThousand(int number) {
        String soFar;

        if (number % 100 < 20) {//không phải là 00
            soFar = numNames[number % 100];
            number /= 100;
        } else {
            //đọc số hàng đơn vị
            soFar = numNames[number % 10];
            number /= 10;

            //ten hang chục
            soFar = tensNames[number % 10] + soFar;
            number /= 10;
        }
        if (number == 0)
            return soFar;
        return numNames[number] + " trăm" + soFar;
    }




    public static String convert(BigDecimal number) {
        // 0 to 999 999 999 999
        if (number.compareTo(BigDecimal.ZERO) == 0) {
            return "zero";
        }

        String snumber = number.toString();

        // pad with "0"
        String mask = "000000000000";
        DecimalFormat df = new DecimalFormat(mask);
        snumber = df.format(number);

        // XXXnnnnnnnnn
        int billions = Integer.parseInt(snumber.substring(0, 3));
        // nnnXXXnnnnnn
        int millions = Integer.parseInt(snumber.substring(3, 6));
        // nnnnnnXXXnnn
        int hundredThousands = Integer.parseInt(snumber.substring(6, 9));
        // nnnnnnnnnXXX
        int thousands = Integer.parseInt(snumber.substring(9, 12));

        String tradBillions;
        switch (billions) {
            case 0:
                tradBillions = "";
                break;
            case 1:
                tradBillions = convertLessThanOneThousand(billions) + " tỷ ";
                break;
            default:
                tradBillions = convertLessThanOneThousand(billions) + " tỷ ";
        }
        String result = tradBillions;

        String tradMillions;
        switch (millions) {
            case 0:
                tradMillions = "";
                break;
            case 1:
                tradMillions = convertLessThanOneThousand(millions) + " triệu ";
                break;
            default:
                tradMillions = convertLessThanOneThousand(millions) + " triệu ";
        }
        result = result + tradMillions;

        String tradHundredThousands;
        switch (hundredThousands) {
            case 0:
                tradHundredThousands = "";
                break;
            case 1:
                tradHundredThousands = "một trăm ";
                break;
            default:
                tradHundredThousands = convertLessThanOneThousand(hundredThousands) + " ngàn ";
        }
        result = result + tradHundredThousands;

        String tradThousand;
        tradThousand = convertLessThanOneThousand(thousands);
        result = result + tradThousand;

        // remove extra spaces!
        return result.replaceAll("^\\s+", "").replaceAll("\\b\\s{2,}\\b", " ");
    }
}

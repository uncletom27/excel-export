package com.meituan.show.settlement.export;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.ServerSocket;
import java.net.Socket;
import java.util.Date;

public class Test {
    public static void main(String[] args) throws IOException {
        Server server = new Server();
        server.start();
        Socket s = new Socket("localhost", 8090);
        OutputStream fos = s.getOutputStream();
        Excel exel = new ExcelImpl(fos);
        exel.beginNewSheet("表");
        exel.addTitle(Test.class);
        for (int i = 0; i < 1000000000; i++) {
            TestModel t = new TestModel();
            exel.addRow(t);
        }
        exel.endSheet();
        exel.finish();
        s.close();
        System.err.println("finished");
    }
    
    private static class Server extends Thread {

        @Override
        public void run() {
            long l = 0;
            ServerSocket ss;
            try (FileOutputStream fos = new FileOutputStream("test.xlsx")){
                ss = new ServerSocket(8090);
                Socket s = ss.accept();
                InputStream inputStream = s.getInputStream();
                byte[] buff = new byte[1024];
                int read = -1;
                while((read = inputStream.read(buff)) != -1){
                    l = l+ read;
                    System.err.println(l);
                    fos.write(buff, 0, read);
                }
                
            } catch (IOException e) {
                e.printStackTrace();
            }
            
        }
        
    }
    
    private static class TestModel {
        @Cell(order = 1)
        Date d = new Date();
        @Cell(order = 2)
        private String s = "大法师打发是发送到发送到发送到发送到发你好阿萨斯的发送";
        @Cell(order = 2)
        private long l = 123l;
        @Cell(order = 3)
        private boolean b = true;
        @Cell(order = 4)
        private  float f = 123.2f;
        @Cell(order = 4)
        private double dd = 123.2d;
        
    }
}

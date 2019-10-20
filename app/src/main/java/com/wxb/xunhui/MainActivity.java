package com.wxb.xunhui;
import android.Manifest;
import android.content.Context;
import android.content.Intent;
import android.graphics.Bitmap;
import android.os.Bundle;
import android.os.Environment;
import android.provider.MediaStore;
import android.support.v7.app.AppCompatActivity;
import android.text.InputType;
import android.view.View;
import android.widget.Button;
import android.widget.EditText;
import android.widget.ImageView;
import android.widget.TextView;
import android.widget.Toast;
import wxbzxing.activity.CaptureActivity;
import wxbzxing.common.BitmapUtils;

import com.qmuiteam.qmui.widget.dialog.QMUIDialog;
import com.qmuiteam.qmui.widget.dialog.QMUIDialogAction;
import com.yanzhenjie.permission.AndPermission;
import com.yanzhenjie.permission.PermissionNo;
import com.yanzhenjie.permission.PermissionYes;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Time;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

public class MainActivity extends AppCompatActivity implements View.OnClickListener {

    private Button mBtn1;
    private EditText mEt;
    private Button mBtn2;
    private Button mBtn3;
    private ImageView mImage;
    private final static int REQ_CODE = 1028;
    private Context mContext;
    private TextView mTvResult;
    String url1 = Environment.getExternalStorageDirectory().getPath()+"/Test/巡回记录表.xls"; //目标文件
    File file;
    String url2;
    Button writecontes;
    long flag=0;
    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
       if(AndPermission.hasPermission(this,Manifest.permission.WRITE_EXTERNAL_STORAGE,Manifest.permission.READ_EXTERNAL_STORAGE)){
        }else{
            AndPermission.with(this)
                    .requestCode(100)
                    .permission(Manifest.permission.WRITE_EXTERNAL_STORAGE,
                                Manifest.permission.READ_EXTERNAL_STORAGE,
                                Manifest.permission.CAMERA
                            )
                    .send();

        }
        initView();
        file=new File(Environment.getExternalStorageDirectory().getPath()+"/Test");
        if(!file.exists()){
        file.mkdir();
        }
        file=new File(Environment.getExternalStorageDirectory().getPath()+"/Test/巡回记录表.xls");
        if(!file.exists()){
            try {
            //file.createNewFile();注意，在建立excel后，要建立相应的表格即sheet，后面才能getSheet，不然写不进去，
               // 之前能建立Excel，但是写不进去，其实是ile.createNewFile()只新建了文件，没有新建sheet,
                // 在createExcel(file)函数中，新建了Excel，还新建了shette,所以可行了。
                createExcel(file);//在createExcel里新建excel，并新建sheet
            }
            catch (Exception e)
            {

            }
        }
        //creatExcelNow("a");
        mContext = this;
    }
    public void createExcel(File file) {
        WritableWorkbook wookbook = null;
        try {
            file.createNewFile();
            //没必要用文件流
            wookbook = Workbook.createWorkbook(file);
            WritableSheet sheet = wookbook.createSheet("巡回记录表", 0);
            wookbook.write();
            wookbook.close();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {

        }

    }
    private void initView() {
        mBtn1 = (Button) findViewById(R.id.btn1);
        mBtn1.setOnClickListener(this);
        mEt = (EditText) findViewById(R.id.et);
        mBtn2 = (Button) findViewById(R.id.btn2);
        mBtn2.setOnClickListener(this);
        mImage = (ImageView) findViewById(R.id.image);
        mImage.setOnClickListener(this);
        mTvResult = (TextView) findViewById(R.id.tv_result);
        mTvResult.setOnClickListener(this);
       findViewById(R.id.tijiao).setOnClickListener(this);
       findViewById(R.id.baocun).setOnClickListener(this);
    }

    @Override
    public void onClick(View v) {
        switch (v.getId()) {
            case R.id.btn1:
               //startActivity(new Intent(MainActivity.this, CaptureActivity.class));
                Intent intent = new Intent(mContext, CaptureActivity.class);
                startActivityForResult(intent, REQ_CODE);
                break;
            case R.id.btn2:
                mImage.setVisibility(View.VISIBLE);
                //隐藏扫码结果view
                mTvResult.setVisibility(View.GONE);

                String content = mEt.getText().toString().trim();
                Bitmap bitmap = null;
                try {
                    bitmap = BitmapUtils.create2DCode(content);//根据内容生成二维码
                    mTvResult.setVisibility(View.GONE);
                    mImage.setImageBitmap(bitmap);
                } catch (Exception e) {
                    e.printStackTrace();
                }
                break;
            case R.id.baocun:
                showReNamePointDialog();
                break;
            case R.id.tijiao:

                        File file=new File(Environment.getExternalStorageDirectory().getPath()+"/Test/巡回记录表.xls");
                        if(file.exists()){
                            FileInputStream in = null;
                            FileOutputStream out=null;
                            try {
                                in = new FileInputStream(new File(url1));
                                out = new FileOutputStream(new File(url2));
                                byte[] buff = new byte[10240]; //限制大小
                                int n = 0;
                                while ((n = in.read(buff)) != -1) {
                                    out.write(buff, 0, n);
                                 }
                                out.flush();
                                in.close();
                                out.close();
                            } catch (Exception e) {
                                e.printStackTrace();
                                Toast.makeText(MainActivity.this,"退出",Toast.LENGTH_SHORT).show();
                            }
                        }
                break;
        }
    }

    @Override
    protected void onActivityResult(int requestCode, int resultCode, Intent data) {
        super.onActivityResult(requestCode, resultCode, data);
        if (requestCode == REQ_CODE) {
            mImage.setVisibility(View.GONE);
            mTvResult.setVisibility(View.VISIBLE);
            String result = data.getStringExtra(CaptureActivity.SCAN_QRCODE_RESULT);
            String strExit="您已取消";
            if(result.equals(strExit)) {
                mTvResult.setText(result);
                showToast( result);
            }
            else{
                mTvResult.setText("扫码结果：" + result);
                showToast("扫码结果：" + result);
                addResult(mTvResult.getText().toString(), file);
            }
        }


    }

    @PermissionYes(100)
    private void getPermission(List<String> grantedPermissions) {
        Toast.makeText(MainActivity.this, "接受权限", Toast.LENGTH_SHORT).show();
    }
    @PermissionNo(100)
    private void refusePermission(List<String> grantedPermissions) {
        Toast.makeText(MainActivity.this, "拒接了权限", Toast.LENGTH_SHORT).show();
    }

    private void showToast(String msg) {
        Toast.makeText(mContext, "" + msg, Toast.LENGTH_SHORT).show();
    }

public  void creatExcelNow(String st){
    SimpleDateFormat dateFormat=new SimpleDateFormat("YYY-MM-dd");
    String nameDate=dateFormat.format(new Date());
    Date date=new Date();
    long hour=date.getHours();
    String hourString;
    if(hour>=0&&hour<=9)
    {
        hourString="夜班";
    }else if(hour>9&&hour<=17){
        hourString="白班";
    }else{
        hourString="中班";
    }
    //url2=Environment.getExternalStorageDirectory().getPath()+"/Test/巡回记录表"+nameDate+hourString+".xls";
    url2=Environment.getExternalStorageDirectory().getPath()+"/Test/巡回记录表"+nameDate+hourString+"("+st+")"+".xls";
    createExcel(new File(url2));
}



    public void addResult(String str,File file) {
        Workbook original = null;
        WritableWorkbook workbook = null;
        try {//  如果是想要修改一个已存在的excel工作簿，则需要先获得它的原始工作簿，再创建一个可读写的副本：
            original = Workbook.getWorkbook(file);
            workbook = Workbook.createWorkbook(file, original);
            WritableSheet sheet = workbook.getSheet(0);
            int row = sheet.getRows();//cpl，列，row,行
            Label label = new Label(0, row, str);
            sheet.addCell(label);
            workbook.write();
        } catch (Exception e) {
        } finally {
            if (original != null) {
                original.close();
            }
            if (workbook != null) {
                try {
                    workbook.close();
                } catch (IOException e) {
                    e.printStackTrace();
                } catch (WriteException e) {
                    e.printStackTrace();
                }
            }

        }
    }
    //重命名节点
    public void showReNamePointDialog() {
        final QMUIDialog.EditTextDialogBuilder builder = new QMUIDialog.EditTextDialogBuilder(MainActivity.this);
       //使用QMUIDialog要先该主题style name="AppTheme" parent="QMUI.Compat"
        builder.setTitle("保存为...")
                .setInputType(InputType.TYPE_CLASS_TEXT)
                .setPlaceholder("在此输入新名称")
                .addAction("取消", new QMUIDialogAction.ActionListener() {
                    @Override
                    public void onClick(QMUIDialog dialog, int index) {
                        dialog.dismiss();
                    }
                })
               .addAction("确认", new QMUIDialogAction.ActionListener() {
                    @Override
                    public void onClick(QMUIDialog dialog, int index) {
                        String newName = builder.getEditText().getText().toString().trim();
                        if (newName.isEmpty()) {
                            Toast.makeText(MainActivity.this, "文件名不合法，请重新输入", Toast.LENGTH_SHORT).show();

                        } else {
                            creatExcelNow(newName);
                            flag=1;
                        }
                        dialog.dismiss();
                    }
                })
                .show();
            }


}



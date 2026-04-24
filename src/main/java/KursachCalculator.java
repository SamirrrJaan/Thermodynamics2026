import com.hummeling.if97.IF97;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

public class KursachCalculator {

    private String name;

    //Начальные параметры общие
    private double x0 = 1;
    private double xc = 0.98;
    private double delta_tпп = 20;
    private double delta_t = 5;
    private double pкн = 1.5;
    //КПД:
    private double eta_oi_ЧВД = 0.88;
    private double eta_oi_ЧНД = 0.85;
    private double eta_oi_ПН = 0.74;
    private double eta_oi_КН = 0.76;
    private double eta_пг = 0.97;
    private double eta_р = 0.95;
    private double eta_мг = 0.98;
    //Начальные параметры по варианту.
    private double Qp;
    private double p0;
    private double pпп;
    private double pk;
    private double tпв;

    public KursachCalculator(String name, double Qp, double p0, double pпп, double pk, double tпв) throws IOException {
        this.Qp = Qp;
        this.p0 = p0;
        this.pпп = pпп;
        this.pk = pk/1000;
        this.tпв = tпв;
        //т.к. внутри программы кельвины то
        this.tпв += 273.15;
        this.name = name;

        doCalculations();
        createTable();

    }

    ArrayList<WaterState> wsList;

    private void doCalculations() throws IOException {
        wsList = new ArrayList<>();
        IF97 if97 = new IF97();

        WaterState ws0 = new WaterState(
                "Точка 0",
                p0, //p
                -1, //t
                -1, //v
                -1, //h
                -1, //s
                1 //x
        );
        wsList.add(ws0);

        WaterState ws1 = new WaterState(
                "Точка 1",
                p0, //p
                -1, //t
                -1, //v
                -1, //h
                -1, //s
                1 //x
        );
        wsList.add(ws1);

        WaterState ws2t = new WaterState(
                "Точка 2t",
                pпп, //p
                -1, //t
                -1, //v
                -1, //h
                ws1.getS(), //s
                -1 //x
        );
        wsList.add(ws2t);

        WaterState ws2 = new WaterState(
                "Точка 2",
                pпп, //p
                -1, //t
                -1, //v
                (ws1.getH() - (ws1.getH() - ws2t.getH())*eta_oi_ЧВД ), //h
                -1, //s
                -1 //x
        );
        wsList.add(ws2);

        WaterState ws3 = new WaterState(
                "Точка 3",
                pпп, //p
                -1, //t
                -1, //v
                -1, //h
                -1, //s
                xc //x
        );
        wsList.add(ws3);

        WaterState ws4 = new WaterState(
                "Точка 4",
                pпп, //p
                (ws1.getT() - delta_tпп), //t
                -1, //v
                -1, //h
                -1, //s
                -1 //x
        );
        wsList.add(ws4);

        WaterState ws5t = new WaterState(
                "Точка 5t",
                pk, //p
                -1, //t
                -1, //v
                -1, //h
                ws4.getS(), //s
                -1 //x
        );
        wsList.add(ws5t);

        WaterState ws5 = new WaterState(
                "Точка 5",
                pk, //p
                -1, //t
                -1, //v
                (ws4.getH() - (ws4.getH() - ws5t.getH())*eta_oi_ЧНД), //h
                -1, //s
                -1 //x
        );
        wsList.add(ws5);

        WaterState ws6 = new WaterState(
                "Точка 6",
                pk, //p
                -1, //t
                -1, //v
                -1, //h
                -1, //s
                0 //x
        );
        wsList.add(ws6);

        WaterState ws7t = new WaterState(
                "Точка 5t",
                pкн, //p
                -1, //t
                -1, //v
                -1, //h
                ws6.getS(), //s
                -1 //x
        );
        wsList.add(ws7t);

        WaterState ws7 = new WaterState(
                "Точка 7",
                pкн, //p
                -1, //t
                -1, //v
                (ws6.getH() + (ws7t.getH() - ws6.getH())/eta_oi_КН ), //h
                -1, //s
                -1 //x
        );
        wsList.add(ws7);

        WaterState ws8 = new WaterState(
                "Точка 8",
                ws7.getP(), //p
                (ws7.getT() + (tпв - ws7.getT())/3), //t
                -1, //v
                -1, //h
                -1, //s
                -1 //x
        );
        wsList.add(ws8);

        WaterState ws9 = new WaterState(
                "Точка 9",
                ws8.getP(), //p
                (ws8.getT() + (tпв - ws7.getT())/3), //t
                -1, //v
                -1, //h
                -1, //s
                -1 //x
        );
        wsList.add(ws9);

        WaterState ws10t = new WaterState(
                "Точка 10t",
                ws0.getP(), //p
                -1, //t
                -1, //v
                -1, //h
                ws9.getS(), //s
                -1 //x
        );
        wsList.add(ws10t);

        WaterState ws10 = new WaterState(
                "Точка 10",
                ws0.getP(), //p
                -1, //t
                -1, //v
                (ws9.getH() + (ws10t.getH() - ws9.getH())/eta_oi_ПН), //h
                -1, //s
                -1 //x
        );
        wsList.add(ws10);

        WaterState wsПВ = new WaterState(
                "Точка пв",
                ws0.getP(), //p
                tпв, //t
                -1, //v
                -1, //h
                -1, //s
                -1 //x
        );
        wsList.add(wsПВ);

        WaterState ws11 = new WaterState(
                "Точка 11",
                ws0.getP(), //p
                -1, //t
                -1, //v
                -1, //h
                -1, //s
                1 //x
        );
        wsList.add(ws11);

        WaterState wsДР_С = new WaterState(
                "Точка 11",
                ws2.getP(), //p
                -1, //t
                -1, //v
                -1, //h
                -1, //s
                0 //x
        );
        wsList.add(wsДР_С);

        WaterState wsДР_ПП = new WaterState(
                "Точка 11",
                ws11.getP(), //p
                -1, //t
                -1, //v
                -1, //h
                -1, //s
                0 //x
        );
        wsList.add(wsДР_ПП);

        WaterState wsОТ1t = new WaterState(
                "Точка от1t",
                if97.saturationPressureT(tпв + delta_t), //p
                -1, //t
                -1, //v
                -1, //h
                ws1.getS(), //s
                -1 //x
        );
        wsList.add(wsОТ1t);

        WaterState wsОТ1 = new WaterState(
                "Точка от1",
                wsОТ1t.getP(), //p
                -1, //t
                -1, //v
                (ws1.getH() - (ws1.getH() - wsОТ1t.getH())*eta_oi_ЧВД), //h
                -1, //s
                -1 //x
        );
        wsList.add(wsОТ1);

        WaterState wsДР1 = new WaterState(
                "Точка др1",
                wsОТ1.getP(), //p
                -1, //t
                -1, //v
                -1, //h
                -1, //s
                0 //x
        );
        wsList.add(wsДР1);

        WaterState wsОТ2t = new WaterState(
                "Точка от2t",
                if97.saturationPressureT(ws9.getT() + delta_t), //p
                -1, //t
                -1, //v
                -1, //h
                ws4.getS(), //s
                -1 //x
        );
        wsList.add(wsОТ2t);

        WaterState wsОТ2 = new WaterState(
                "Точка от2",
                wsОТ2t.getP(), //p
                -1, //t
                -1, //v
                (ws4.getH() - (ws4.getH() - wsОТ2t.getH())*eta_oi_ЧНД), //h
                -1, //s
                -1 //x
        );
        wsList.add(wsОТ2);

        WaterState wsОТ3t = new WaterState(
                "Точка от3t",
                if97.saturationPressureT(ws8.getT() + delta_t), //p
                -1, //t
                -1, //v
                -1, //h
                ws4.getS(), //s
                -1 //x
        );
        wsList.add(wsОТ3t);

        WaterState wsОТ3 = new WaterState(
                "Точка от3",
                wsОТ3t.getP(), //p
                -1, //t
                -1, //v
                (ws4.getH() - (ws4.getH() - wsОТ3t.getH())*eta_oi_ЧНД), //h
                -1, //s
                -1 //x
        );
        wsList.add(wsОТ3);

        WaterState wsДР3 = new WaterState(
                "Точка др3",
                wsОТ3.getP(), //p
                -1, //t
                -1, //v
                -1, //h
                -1, //s
                0 //x
        );
        wsList.add(wsДР3);

        for(WaterState ws: wsList) {
            ws.printParameters();
        }

    }

    private void createTable() throws IOException {
        //Blank Document
        XWPFDocument document = new XWPFDocument();

        //Write the Document in file system
        FileOutputStream out = new FileOutputStream(new File(name + ".docx"));

        //create table
        XWPFTable table = document.createTable();

        //create first row
        XWPFTableRow tableRowOne = table.getRow(0);
        tableRowOne.getCell(0).setText("Характерная точка установки");
        tableRowOne.addNewTableCell().setText("P, МПа");
        tableRowOne.addNewTableCell().setText("t, 'C");
        tableRowOne.addNewTableCell().setText("v, м3/кг");
        tableRowOne.addNewTableCell().setText("h, кДж/кг");
        tableRowOne.addNewTableCell().setText("s, кДж/(кг*К)");
        tableRowOne.addNewTableCell().setText("x -");
        tableRowOne.addNewTableCell().setText("alpha");

        //create second row
        String[] params;
        XWPFTableRow[] rows = new XWPFTableRow[wsList.size()];
        for(int i = 0; i < rows.length; i++) {
            rows[i] = table.createRow();
        }
        for(int i = 0; i < rows.length; i++) {
            params = wsList.get(i).getParameters();
            for(int j = 0; j < params.length; j++) {
                StringBuilder sb = new StringBuilder(params[j]);
                for(int k = 0; k < sb.length(); k ++) {
                    if(sb.charAt(k) == '.') {
                        sb.replace(k, k+1, ",");
                    }
                }
                rows[i].getCell(j).setText(sb.toString());
            }
        }

        document.write(out);
        out.close();
        document.close();
        System.out.println(name + ".docx written successully");
    }

}

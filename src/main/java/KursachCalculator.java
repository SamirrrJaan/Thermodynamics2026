import com.hummeling.if97.IF97;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

public class KursachCalculator {

    private final String name;

    //Начальные параметры общие
    private final double x0 = 1;
    private final double xc = 0.98;
    private final double delta_tпп = 20;
    private final double delta_t = 5;
    private final double pкн = 1.5;
    //КПД:
    private final double eta_oi_ЧВД = 0.88;
    private final double eta_oi_ЧНД = 0.85;
    private final double eta_oi_ПН = 0.74;
    private final double eta_oi_КН = 0.76;
    private final double eta_пг = 0.97;
    private final double eta_р = 0.95;
    private final double eta_мг = 0.98;
    //Начальные параметры по варианту.
    private final double Qp;
    private final double p0;
    private final double pпп;
    private final double pk;
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

    private ArrayList<WaterState> wsList;
    double l_pt, l_pt_ЧВД, l_pt_ЧНД;
    double l_n, l_n_pn, l_n_kn;
    double l_c, q1, eta_i, d0, q0, D0, N_e;

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

        WaterState ws_pv = new WaterState(
                "Точка пв",
                ws0.getP(), //p
                tпв, //t
                -1, //v
                -1, //h
                -1, //s
                -1 //x
        );
        wsList.add(ws_pv);

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

        WaterState ws_dr_s = new WaterState(
                "Точка др_с",
                ws2.getP(), //p
                -1, //t
                -1, //v
                -1, //h
                -1, //s
                0 //x
        );
        wsList.add(ws_dr_s);

        WaterState ws_dr_pp = new WaterState(
                "Точка др_пп",
                ws11.getP(), //p
                -1, //t
                -1, //v
                -1, //h
                -1, //s
                0 //x
        );
        wsList.add(ws_dr_pp);

        WaterState ws_ot_1t = new WaterState(
                "Точка от1t",
                if97.saturationPressureT(tпв + delta_t), //p
                -1, //t
                -1, //v
                -1, //h
                ws1.getS(), //s
                -1 //x
        );
        wsList.add(ws_ot_1t);

        WaterState ws_ot_1 = new WaterState(
                "Точка от1",
                ws_ot_1t.getP(), //p
                -1, //t
                -1, //v
                (ws1.getH() - (ws1.getH() - ws_ot_1t.getH())*eta_oi_ЧВД), //h
                -1, //s
                -1 //x
        );
        wsList.add(ws_ot_1);

        WaterState ws_dr_1 = new WaterState(
                "Точка др1",
                ws_ot_1.getP(), //p
                -1, //t
                -1, //v
                -1, //h
                -1, //s
                0 //x
        );
        wsList.add(ws_dr_1);

        WaterState ws_ot_2t = new WaterState(
                "Точка от2t",
                if97.saturationPressureT(ws9.getT() + delta_t), //p
                -1, //t
                -1, //v
                -1, //h
                ws4.getS(), //s
                -1 //x
        );
        wsList.add(ws_ot_2t);

        WaterState ws_ot_2 = new WaterState(
                "Точка от2",
                ws_ot_2t.getP(), //p
                -1, //t
                -1, //v
                (ws4.getH() - (ws4.getH() - ws_ot_2t.getH())*eta_oi_ЧНД), //h
                -1, //s
                -1 //x
        );
        wsList.add(ws_ot_2);

        WaterState ws_ot_3t = new WaterState(
                "Точка от3t",
                if97.saturationPressureT(ws8.getT() + delta_t), //p
                -1, //t
                -1, //v
                -1, //h
                ws4.getS(), //s
                -1 //x
        );
        wsList.add(ws_ot_3t);

        WaterState ws_ot_3 = new WaterState(
                "Точка от3",
                ws_ot_3t.getP(), //p
                -1, //t
                -1, //v
                (ws4.getH() - (ws4.getH() - ws_ot_3t.getH())*eta_oi_ЧНД), //h
                -1, //s
                -1 //x
        );
        wsList.add(ws_ot_3);

        WaterState ws_dr_3 = new WaterState(
                "Точка др3",
                ws_ot_3.getP(), //p
                -1, //t
                -1, //v
                -1, //h
                -1, //s
                0 //x
        );
        wsList.add(ws_dr_3);

        //Расчёт отборов
        double alpha_ot_1;
        double alpha_ot_2;
        double alpha_ot_3;
        double alpha_dr_s ;
        double alpha_dr_pp;

        alpha_dr_s = ( ( H(ws_ot_1) - H(ws_dr_1) )/( H(ws_pv) - H(ws10) ) -1 ) /
                /**/( ( (H(ws_dr_s) - H(ws3))/( H(ws3) - H(ws2) ) ) * ( ((H(ws3) - H(ws4))/(H(ws11) - H(ws_dr_pp))) - ((H(ws_ot_1) - H(ws_dr_1) )/(H(ws_pv) - H(ws10))) ) );/**/
        System.out.println("alpha_dr_s " + alpha_dr_s);
        ws_dr_s.setSelection(alpha_dr_s);

        alpha_ot_1 = 1 + ( H(ws_dr_s) - H(ws3) ) / (H(ws3) - H(ws2)) * alpha_dr_s;
        System.out.println("alpha_ot1 " + alpha_ot_1);
        ws_ot_1.setSelection(alpha_ot_1);

        alpha_dr_pp = ( H(ws_ot_1) - H(ws_dr_1) )/( H(ws_pv) - H(ws10) ) * alpha_ot_1 - 1;
        System.out.println("alpha_dr_pp " + alpha_dr_pp);
        ws_dr_pp.setSelection(alpha_dr_pp);

        alpha_ot_2 = ((1 + alpha_dr_pp) * H(ws9) - alpha_ot_1 * H(ws_dr_1) - alpha_dr_pp * H(ws_dr_pp) - H(ws8) + alpha_ot_1 * H(ws8)) / (H(ws_ot_2) - H(ws8));
        System.out.println("alpha_ot2 " + alpha_ot_2);
        ws_ot_2.setSelection(alpha_ot_2);

        alpha_ot_3 = ( (1 - alpha_ot_1 - alpha_ot_2) * (H(ws8) - H(ws7)) - alpha_dr_s * (H(ws_dr_s) - H(ws_dr_3))) / (H(ws_ot_3) - H(ws_dr_3));
        System.out.println("alpha_ot3 " + alpha_ot_3);
        ws_ot_3.setSelection(alpha_ot_3);

        ws1.setSelection(1);
        ws2.setSelection(1-alpha_ot_1);
        ws3.setSelection(1 - alpha_ot_1 - alpha_dr_s);
        ws4.setSelection(ws3.getSelection());
        ws5.setSelection(ws4.getSelection() - alpha_ot_2 - alpha_ot_3);
        ws6.setSelection(1 - alpha_ot_1 - alpha_ot_2);
        ws7.setSelection(ws6.getSelection());
        ws8.setSelection(ws7.getSelection());
        ws9.setSelection(1 + alpha_dr_pp);
        ws10.setSelection(ws9.getSelection());
        ws_pv.setSelection(ws10.getSelection());
        ws11.setSelection(alpha_dr_pp);
        ws0.setSelection(ws_pv.getSelection());
        ws_dr_1.setSelection(alpha_ot_1);
        ws_dr_3.setSelection(alpha_ot_3 + alpha_dr_s);

        //Прочие характеристики
        l_pt_ЧВД = (H(ws1) - H(ws2)) - alpha_ot_1 * (H(ws_ot_1) - H(ws2));
        l_pt_ЧНД  = (1 - alpha_ot_1 - alpha_dr_s) * (H(ws4) - H(ws5)) -
                alpha_ot_2 * (H(ws_ot_2) - H(ws5)) -
                alpha_ot_3 * (H(ws_ot_3) - H(ws5));
        l_pt = l_pt_ЧВД + l_pt_ЧНД;

        l_n_pn = ws10.getSelection() * (H(ws10) - H(ws9));
        l_n_kn = ws6.getSelection() * (H(ws7) - H(ws6));
        l_n = l_n_kn + l_n_pn;

        l_c = l_pt - l_n;

        q1 = ws0.getSelection() * (H(ws1) - H(ws_pv));
        eta_i = l_c / q1;
        d0 = 3600/l_c;
        q0 = 3600/eta_i;
        D0 = Qp * 1000 / (q1 * eta_р * eta_мг);
        N_e = D0 * l_c * eta_мг / 1000;

//        for(WaterState ws: wsList) {
//            ws.printParameters();
//        }

    }

    private double H(WaterState ws) {
        return ws.getH();
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
        document.createParagraph().setPageBreak(false);
        XWPFRun run = document.createParagraph().createRun();
        run.setText("l_pt_ЧВД: " + l_pt_ЧВД);
        run.addBreak();
        run.setText("l_pt_ЧНД: " + l_pt_ЧНД);
        run.addBreak();
        run.setText("l_pt: " + l_pt);
        run.addBreak();
        run.setText("l_n_kn: " + l_n_kn);
        run.addBreak();
        run.setText("l_n_pn: " + l_n_pn);
        run.addBreak();
        run.setText("l_n: " + l_n);
        run.addBreak();
        run.setText("l_c: " + l_c);
        run.addBreak();
        run.setText("q1: " + q1);
        run.addBreak();
        run.setText("eta_i: " + eta_i);
        run.addBreak();
        run.setText("d0: " + d0);
        run.addBreak();
        run.setText("q0: " + q0);
        run.addBreak();
        run.setText("D0: " + D0);
        run.addBreak();
        run.setText("N_e: " + N_e);
        run.addBreak();

        document.write(out);
        out.close();
        document.close();
        System.out.println(name + ".docx written successfully");
    }

}

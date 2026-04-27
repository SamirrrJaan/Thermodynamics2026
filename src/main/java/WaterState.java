import com.hummeling.if97.IF97;

public class WaterState {
    private IF97 if97 = new IF97();

    private String name;
    private double pressure;
    private double temperature;
    private double volume;
    private double enthalpy;
    private double entropy;
    private double dryness;
    private double selection;

    //params = { 12, p
    //           12, t
    //           12, v
    //           12, h
    //           12, s
    //           12, x
    //           12  alpha
    //}
    public WaterState
    (
            String name,
            double pressure,
            double temperature,
            double volume,
            double enthalpy,
            double entropy,
            double dryness
    )
    {
        this.name = name;
        this.pressure = pressure;
        this.temperature = temperature;
        this.volume = volume;
        this.enthalpy = enthalpy;
        this.entropy = entropy;
        this.dryness = dryness;
        fillOtherParameters();
        selection = -100500;
    }

    //Возможные комбинации параметров входных:
    // p x
    // p s
    // p h
    // p t

    private void fillOtherParameters() {
        if(entropy != -1) {
            fillPS();
        } else if(enthalpy != -1) {
            fillPH();
        } else if(temperature != -1) {
            fillPT();
        } else if(dryness != -1) {
            fillPX();
        }
        roundEverything();
    }

    private void fillPS() {
        temperature = if97.temperaturePS(pressure, entropy);
        volume = if97.specificVolumePS(pressure, entropy);
        enthalpy = if97.specificEnthalpyPS(pressure, entropy);
        dryness = if97.vapourFractionPS(pressure, entropy);
    }

    private void fillPH() {
        temperature = if97.temperaturePH(pressure, enthalpy);
        volume = if97.specificVolumePH(pressure, enthalpy);
        entropy = if97.specificEntropyPH(pressure, enthalpy);
        dryness = if97.vapourFractionPH(pressure, enthalpy);

    }
    private void fillPT() {
        entropy = if97.specificEntropyPT(pressure, temperature);
        volume = if97.specificVolumePS(pressure, entropy);
        enthalpy = if97.specificEnthalpyPT(pressure, temperature);
        dryness = if97.vapourFractionPS(pressure, entropy);
    }

    private void fillPX() {
        enthalpy = if97.specificEnthalpyPX(pressure, dryness);
        temperature = if97.temperaturePH(pressure, enthalpy);
        volume = if97.specificVolumePX(pressure, dryness);
        entropy = if97.specificEntropyPX(pressure, dryness);
    }

    private void roundEverything() {
        //Сколько знаков после запятой оставлять:
        int howMuchDigits = 3;
        howMuchDigits = (int) Math.pow(10, howMuchDigits);
        pressure = (double) Math.round(pressure * howMuchDigits) / howMuchDigits;
        temperature = (double) Math.round(temperature * howMuchDigits) / howMuchDigits;
        volume = (double) Math.round(volume * howMuchDigits) / howMuchDigits;
        enthalpy = (double) Math.round(enthalpy * howMuchDigits) / howMuchDigits;
        entropy = (double) Math.round(entropy * howMuchDigits) / howMuchDigits;
        dryness = (double) Math.round(dryness * howMuchDigits) / howMuchDigits;
        selection = (double) Math.round(selection * howMuchDigits) / howMuchDigits;
    }

    public void printParameters() {
        System.out.println(name);
        System.out.println("P = " + pressure);
        System.out.println("t = " + temperature);
        System.out.println("v = " + volume);
        System.out.println("h = " + enthalpy);
        System.out.println("s = " + entropy);
        System.out.println("x = " + dryness);
        System.out.println("a = " + selection);
    }

    public String[] getParameters() {
        temperature -= 273.15;
        temperature = (double) Math.round(temperature * 3) / 3;
        String drynessStr = "" + dryness;
        String selectionStr = "" + selection;
        if(dryness > 1 || dryness < 0) {
            drynessStr = "-";
        }
        if(selection == -100500) {
            selectionStr = "-";
        }
        return new String[]{
                name,
                "" + pressure,
                "" + temperature,
                "" + volume,
                "" + enthalpy,
                "" + entropy,
                drynessStr,
                selectionStr
        };
    }

    public double getP() {
        return pressure;
    }

    public void setP(double pressure) {
        this.pressure = pressure;
    }

    public double getT() {
        return temperature;
    }

    public void setT(double temperature) {
        this.temperature = temperature;
    }

    public double getV() {
        return volume;
    }

    public void setV(double volume) {
        this.volume = volume;
    }

    public double getH() {
        return enthalpy;
    }

    public void setH(double enthalpy) {
        this.enthalpy = enthalpy;
    }

    public double getS() {
        return entropy;
    }

    public void setS(double entropy) {
        this.entropy = entropy;
    }

    public double getX() {
        return dryness;
    }

    public void setX(double dryness) {
        this.dryness = dryness;
    }

    public double getSelection() {
        return selection;
    }

    public void setSelection(double selection) {
        this.selection = selection;
        roundEverything();
    }
}

package Converter;

import java.util.regex.Pattern;

public enum DeviceModel {
    PQ211KA00Z("\\bPQ211KA00Z\\b"),
    HE273ABS0("\\bHE273ABS0\\b"),
    EA645GN17("\\bEA645GN17\\b"),
    HZ531000("\\bHZ531000\\b"),
    SN636X06KE("\\bSN636X06KE\\b"),
    SN73HX42VE("\\bSN73HX42VE\\b"),
    BF525LMS0("\\bBF525LMS0\\b"),
    KI86NNFE0("\\bKI86NNFE0\\b"),
    HB274ABS0("\\bHB274ABS0\\b"),
    HB272ABB0("\\bHB272ABB0\\b"),
    KI86VVFE0("\\bKI86VVFE0\\b"),
    Z8TST096Z1("\\bZ8TST096Z1\\b"),
    EQ212KA01Z("\\bEQ212KA01Z\\b"),
    HB213ABS0("\\bHB213ABS0\\b"),
    ET645BNA1E("\\bET645BNA1E\\b"),
    HZ638200("\\bHZ638200\\b");




    private final Pattern pattern;

    DeviceModel(String regex) {
        this.pattern = Pattern.compile(regex);
    }

    public Pattern getPattern() {
        return pattern;
    }
}

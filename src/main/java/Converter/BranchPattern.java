package Converter;

import java.util.regex.Pattern;

public enum BranchPattern {
    RA("\\b(RA)\\w{0,4}\\b"),
    BV("\\b(BV)\\w{0,4}\\b"),
    JL("\\b(JL)\\w{0,4}\\b"),
    B4("\\b(B4)\\w{0,4}\\b"),
    AX("\\b(AX)\\w{0,4}\\b"),
    NA("\\b(NA)\\w{0,4}\\b"),

    RK("\\b(RK)\\w{0,4}\\b");

    private final Pattern pattern;

    BranchPattern(String regex) {
        this.pattern = Pattern.compile(regex);
    }

    public Pattern getPattern() {
        return pattern;
    }
}

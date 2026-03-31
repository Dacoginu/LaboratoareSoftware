package lab5;

public class AdvancedCalculator extends IntCalculator {

    public AdvancedCalculator(int initialValue) {
        super(initialValue);
    }

    @Override
    public AdvancedCalculator add(int value) {
        super.add(value);
        return this;
    }

    @Override
    public AdvancedCalculator subtract(int value) {
        super.subtract(value);
        return this;
    }

    @Override
    public AdvancedCalculator multiply(int value) {
        super.multiply(value);
        return this;
    }

    public AdvancedCalculator divide(int value) {
        if (value != 0) {
            int res = result() / value;
            clear();
            add(res);
        }
        return this;
    }

    public AdvancedCalculator power(int exponent) {
        int base = result();
        int res = 1;

        for (int i = 0; i < exponent; i++) {
            res *= base;
        }

        clear();
        add(res);
        return this;
    }

    public AdvancedCalculator root(int n) {
        double res = Math.pow(result(), 1.0 / n);
        clear();
        add((int) res);
        return this;
    }
}
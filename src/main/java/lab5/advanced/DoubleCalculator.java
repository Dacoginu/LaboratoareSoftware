package lab5.advanced;

public class DoubleCalculator extends ACalculator {

    @Override
    public void init(Object value) {
        this.state = (Double) value;
    }

    public DoubleCalculator add(Double value) {
        state = (Double) state + value;
        return this;
    }

    public DoubleCalculator subtract(Double value) {
        state = (Double) state - value;
        return this;
    }

    public DoubleCalculator multiply(Double value) {
        state = (Double) state * value;
        return this;
    }
}
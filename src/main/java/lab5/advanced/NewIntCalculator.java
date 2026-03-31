package lab5.advanced;

public class NewIntCalculator extends ACalculator {

    @Override
    public void init(Object value) {
        this.state = (Integer) value;
    }

    public NewIntCalculator add(Integer value) {
        state = (Integer) state + value;
        return this;
    }

    public NewIntCalculator subtract(Integer value) {
        state = (Integer) state - value;
        return this;
    }

    public NewIntCalculator multiply(Integer value) {
        state = (Integer) state * value;
        return this;
    }
}
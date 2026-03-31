package lab5;

public class IntCalculator {
    private int state;



    public IntCalculator add(int value) {
        this.state += value;
        return this;
    }

    public IntCalculator subtract(int value) {
        this.state -= value;
        return this;
    }

    public IntCalculator multiply(int value) {
        this.state *= value;
        return this;
    }

    public int result() {
        return state;
    }

    public IntCalculator clear() {
        this.state = 0;
        return this;
    }

    public IntCalculator(int state) {
        this.state = state;
    }
}
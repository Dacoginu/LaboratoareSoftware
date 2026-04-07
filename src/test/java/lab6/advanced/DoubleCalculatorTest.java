package lab6.advanced;

import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import static org.junit.jupiter.api.Assertions.*;

class DoubleCalculatorTest {

    private DoubleCalculator calc;

    @BeforeEach
    void setUp() {
        calc = new DoubleCalculator();
        calc.init();
    }

    @AfterEach
    void tearDown() {
        calc.clear();
    }

    @Test
    void testAddPositive() {
        calc.add(5.5);
        assertEquals(5.5, calc.result(), 0.0001);
    }

    @Test
    void testAddNegatives() {
        calc.add(-2.5);
        assertEquals(-2.5, calc.result(), 0.0001);
    }

    @Test
    void testSubtractPositives() {
        calc.add(10.5).subtract(3.2);
        assertEquals(7.3, calc.result(), 0.0001);
    }

    @Test
    void testMultiplyPositives() {
        calc.add(2.0).multiply(2.5);
        assertEquals(5.0, calc.result(), 0.0001);
    }

    @Test
    void testDividePositives() {
        calc.add(10.0).divide(2.0);
        assertEquals(5.0, calc.result(), 0.0001);
    }


    @Test
    void testDivideBy0() {
        calc.add(5.0).divide(0.0);
        assertEquals(Double.POSITIVE_INFINITY, calc.result());
    }

    @Test
    void testAddNull() {
        assertThrows(NullPointerException.class, () -> {
            calc.add(null);
        });
    }
}
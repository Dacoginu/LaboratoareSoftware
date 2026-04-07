package lab6.advance;

import lab6.advanced.NewIntCalculator;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import static org.junit.jupiter.api.Assertions.*;

class NewIntCalculatorTest {

    private NewIntCalculator calc;

    @BeforeEach
    void setUp() {
        calc = new NewIntCalculator();
        calc.init();
    }

    @AfterEach
    void tearDown() {
        calc.clear();
    }

    @Test
    void testAddPositive() {
        // Act
        calc.add(5);
        // Assert
        assertEquals(5, calc.result());
    }

    @Test
    void testAddNegatives() {
        calc.add(-3);
        assertEquals(-3, calc.result());
    }

    @Test
    void testSubtractPositives() {
        calc.add(10).subtract(4);
        assertEquals(6, calc.result());
    }

    @Test
    void testSubtractNegatives() {
        calc.subtract(-5); // 0 - (-5) = 5
        assertEquals(5, calc.result());
    }

    @Test
    void testMultiplyPositives() {
        calc.add(3).multiply(4);
        assertEquals(12, calc.result());
    }

    @Test
    void testMultiplyNegatives() {
        calc.add(4).multiply(-2);
        assertEquals(-8, calc.result());
    }

    @Test
    void testMultiplyBy0() {
        calc.add(10).multiply(0);
        assertEquals(0, calc.result());
    }

    @Test
    void testDividePositives() {
        calc.add(10).divide(2);
        assertEquals(5, calc.result());
    }

    @Test
    void testDivideNegatives() {
        calc.add(10).divide(-2);
        assertEquals(-5, calc.result());
    }


    @Test
    void testDivideBy0() {
        calc.add(5);
        assertThrows(ArithmeticException.class, () -> {
            calc.divide(0);
        });
    }

    @Test
    void testAddNull() {
        assertThrows(NullPointerException.class, () -> {
            calc.add(null);
        });
    }
}
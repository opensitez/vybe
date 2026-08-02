// vybe-test: kotlin/kotlin_big_numbers/test_big_decimal_invalid_division
// origin: languages/kotlin/tests/kotlin/test_kotlin_big_numbers.rs

fun main() {
            val a = java.math.BigDecimal("10")
            try {
                a.divide(java.math.BigDecimal("0"))
                println("bad")
            } catch (e: ArithmeticException) {
                println(e::class.simpleName)
            }
        }


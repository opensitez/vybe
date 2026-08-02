// vybe-test: kotlin/advanced_features/test_when_multiple_cases
// origin: languages/kotlin/tests/kotlin/test_advanced_features.rs

fun main() {
            val day = 6
            when (day) {
                1, 2, 3, 4, 5 -> println("weekday")
                6, 7 -> println("weekend")
            }
        }


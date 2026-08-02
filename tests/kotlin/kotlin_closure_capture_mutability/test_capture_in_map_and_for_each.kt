// vybe-test: kotlin/kotlin_closure_capture_mutability/test_capture_in_map_and_for_each
// origin: languages/kotlin/tests/kotlin/test_kotlin_closure_capture_mutability.rs

fun main() {
            var total = 0
            listOf(1, 2, 3).forEach { total += it }
            println(total)
        }


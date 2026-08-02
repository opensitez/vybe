// vybe-test: kotlin/strings/test_string_slice_out_of_bounds_throws
// origin: languages/kotlin/tests/kotlin/test_strings.rs

fun main() {
            val value = "abc"
            try {
                println(value.slice(2..5))
            } catch (e: Exception) {
                println("slice-error")
            }
        }


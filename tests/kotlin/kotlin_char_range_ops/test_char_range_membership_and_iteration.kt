// vybe-test: kotlin/kotlin_char_range_ops/test_char_range_membership_and_iteration
// origin: languages/kotlin/tests/kotlin/test_kotlin_char_range_ops.rs

fun main() {
            val span = 'b'..'e'
            var text = ""
            for (c in span) { text += c }
            println(text)
            println('a' in span)
            println('d' in span)
        }


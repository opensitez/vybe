// vybe-test: kotlin/smart_casts/test_smart_cast_in_while_like_rewrite
// origin: languages/kotlin/tests/kotlin/test_smart_casts.rs

fun toMessage(value: Any): String {
            var cursor: Any = value
            var result = ""
            if (cursor is String) {
                result = cursor + "!"
            }
            if (cursor is String) {
                result += " twice"
            }
            println(result)
            return result
        }

        fun main() {
            println(toMessage("x"))
            println(toMessage(4))
        }


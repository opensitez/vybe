// vybe-test: kotlin/enums/test_enum_iteration_with_for
// origin: languages/kotlin/tests/kotlin/test_enums.rs

enum class Channel {
            RED, GREEN, BLUE
        }

        fun main() {
            var names = ""
            for (item in arrayOf(Channel.RED, Channel.GREEN, Channel.BLUE)) {
                names = names + item + ","
            }
            println(names)
        }


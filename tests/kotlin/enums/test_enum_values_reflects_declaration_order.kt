// vybe-test: kotlin/enums/test_enum_values_reflects_declaration_order
// origin: languages/kotlin/tests/kotlin/test_enums.rs

enum class Color { RED, GREEN, BLUE }

        fun main() {
            val values = Color.values()
            var names = ""
            for (value in values) {
                names += value.name + "|"
            }
            println(names)
            println(values.size)
            println(values[0] == Color.RED)
        }


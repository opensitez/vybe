// vybe-test: kotlin/object_declarations/test_object_expression_can_implement_custom_interface
// origin: languages/kotlin/tests/kotlin/test_object_declarations.rs

interface Printer {
            fun print(): String
        }

        fun main() {
            val value: Printer = object : Printer {
                override fun print(): String = "done"
            }
            println(value.print())
        }


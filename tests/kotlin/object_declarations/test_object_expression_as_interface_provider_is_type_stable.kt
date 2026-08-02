// vybe-test: kotlin/object_declarations/test_object_expression_as_interface_provider_is_type_stable
// origin: languages/kotlin/tests/kotlin/test_object_declarations.rs

interface Printer {
            fun print(): String
        }

        fun makePrinter(prefix: String): Printer {
            return object : Printer {
                override fun print(): String = prefix + "!"
            }
        }

        fun main() {
            val first = makePrinter("a")
            val second = makePrinter("b")
            println(first.print())
            println(second.print())
        }


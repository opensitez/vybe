// vybe-test: kotlin/interfaces/test_interface_default_to_override
// origin: languages/kotlin/tests/kotlin/test_interfaces.rs

interface Source {
            fun value(): Int {
                return 1
            }
        }

        class Provider : Source {
            override fun value(): Int {
                return 2
            }
        }

        fun main() {
            val s: Source = Provider()
            println(s.value())
        }


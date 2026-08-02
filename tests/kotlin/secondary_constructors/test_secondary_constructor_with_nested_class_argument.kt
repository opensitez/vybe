// vybe-test: kotlin/secondary_constructors/test_secondary_constructor_with_nested_class_argument
// origin: languages/kotlin/tests/kotlin/test_secondary_constructors.rs

class Host {
            val name: String
            class Config

            constructor(value: String) {
                this.name = value
            }

            constructor(config: Config, value: String) : this(value) {
                val used = config
                __check((used is Config).toString(), "true")
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Host("root").name).toString(), "root")
            __check((Host(Host.Config(), "inner").name).toString(), "inner")
        }

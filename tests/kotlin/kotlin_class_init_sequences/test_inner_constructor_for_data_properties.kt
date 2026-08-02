// vybe-test: kotlin/kotlin_class_init_sequences/test_inner_constructor_for_data_properties
// origin: languages/kotlin/tests/kotlin/test_kotlin_class_init_sequences.rs

class Profile {
            val name: String
            val suffix: String
            constructor(name: String) {
                this.name = name
                this.suffix = name.takeLast(1)
            }
            constructor(name: String, idx: Int) : this(name) {
                this.suffix = name[idx]
            }

            fun render(): String = name + suffix
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Profile("abc").render()).toString(), "abcc")
            __check((Profile("xyz", 1).render()).toString(), "xyy")
        }

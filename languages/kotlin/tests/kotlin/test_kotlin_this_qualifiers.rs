kotlin_run_cases! {
    test_outer_this_in_instance_method => (r#"
        class Outer {
            val name = "outer"

            inner class Inner {
                fun valueFromOuter(): String {
                    return this@Outer.name
                }
            }
        }

        fun main() {
            val out = Outer().Inner()
            println(out.valueFromOuter())
        }
    "#, vec!["outer"]),
    test_nested_this_in_lambda => (r#"
        class Box {
            val marker = "box"
            inner class Holder {
                fun show(prefix: String): String {
                    val read = this@Holder
                    return prefix + read.parentName()
                }

                fun parentName(): String {
                    return this@Box.marker
                }
            }
        }

        fun main() {
            val h = Box().Holder()
            println(h.show("m="))
        }
    "#, vec!["m=box"]),
    test_this_parameter_capture => (r#"
        class Chain {
            val value = 5

            fun mark(prefix: String): String {
                return this.toStringPrefix(prefix)
            }

            fun toStringPrefix(prefix: String): String {
                return prefix + this.value.toString()
            }
        }

        fun main() {
            println(Chain().mark("v="))
        }
    "#, vec!["v=5"]),
}

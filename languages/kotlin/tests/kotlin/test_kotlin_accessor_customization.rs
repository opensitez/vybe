kotlin_run_cases! {
    test_custom_getter => (r#"
        class Score {
            private var raw = 0

            var value: Int
                get() = raw * 2
                set(v) { raw = if (v < 0) 0 else v }
        }

        fun main() {
            val s = Score()
            s.value = 3
            println(s.value)
            s.value = -4
            println(s.value)
        }
    "#, vec!["6", "0"]),
    test_backed_property_with_side_effect => (r##"
        class Store {
            var marker: String = ""
                set(v) {
                    field = v + "!"
                }
                get() = field + "#"
        }

        fun main() {
            val s = Store()
            s.marker = "go"
            println(s.marker)
        }
    "##, vec!["go!#"]),
    test_property_init_before_setter => (r#"
        class Counter {
            var value = 1
                private set

            fun setPublic(v: Int) {
                value = v
            }
        }

        fun main() {
            val c = Counter()
            c.setPublic(5)
            println(c.value)
        }
    "#, vec!["5"]),
}

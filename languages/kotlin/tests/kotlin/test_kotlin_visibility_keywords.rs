kotlin_run_cases! {
    test_private_and_public_property_access => (r#"
        open class Base {
            private val secret = "secret"
            public val shown = "shown"
            protected open val inherited = "inherited"
        }

        class Child : Base() {
            override val inherited: String = "childInherited"
            fun exposeInherited(): String {
                return inherited
            }
        }

        fun main() {
            val b = Child()
            println(b.shown)
            println(b.exposeInherited())
        }
    "#, vec!["shown", "childInherited"]),
    test_internal_default_visibility => (r#"
        internal const val scope = "module"

        class Counter {
            internal var value = 0
            fun bump(): String {
                value = value + 1
                return scope + value.toString()
            }
        }

        fun main() {
            val c = Counter()
            println(c.bump())
            println(c.bump())
        }
    "#, vec!["module1", "module2"]),
    test_local_visibility_with_private_setter => (r#"
        class Data {
            var value: Int = 0
                private set

            fun assign(next: Int) {
                value = next
            }
        }

        fun main() {
            val d = Data()
            d.assign(3)
            println(d.value)
        }
    "#, vec!["3"]),
}

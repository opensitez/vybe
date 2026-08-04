kotlin_run_test!(
    test_this_in_inner_class_refs_outer,
    r#"
        class Outer(val outerLabel: String) {
            inner class Inner {
                fun full(): String = this@Outer.outerLabel + ":inner"
            }

            fun probe(): String = Inner().full()
        }

        fun main() {
            println(Outer("root").probe())
        }
    "#,
    &["root:inner"]
);

kotlin_run_test!(
    test_extension_receiver_this_in_lambda,
    r#"
        fun String.wrap(): String = this.also { println("start") } + "!"

        fun main() {
            println("ok".wrap())
        }
    "#,
    &["start", "ok!"]
);

kotlin_run_test!(
    test_this_with_label_in_nested_functions,
    r#"
        class Container {
            val name = "container"

            fun make(prefix: String): String {
                fun nested() = this@Container.name + prefix
                return nested()
            }
        }

        fun main() {
            println(Container().make("X"))
        }
    "#,
    &["containerX"]
);

kotlin_run_test!(
    test_extension_function_with_explicit_this_parameter,
    r#"
        class Holder {
            fun Int.addToHolder(): Int = this + 10
        }

        fun main() {
            val h = Holder()
            println(h.run {
                3.addToHolder()
            })
        }
    "#,
    &["13"]
);

kotlin_run_test!(
    test_lambda_with_this_from_scoped_receiver,
    r#"
        data class Box(val value: String)

        fun main() {
            val out = Box("x").run {
                with(this) {
                    println(value)
                    this.value.length
                }
            }
            println(out)
        }
    "#,
    &["x", "1"]
);

kotlin_run_test!(
    test_nested_class_this_qualifier,
    r#"
        class A {
            val name = "A"
            inner class B {
                fun call(): String = this@A.name
            }
        }

        fun main() {
            println(A().B().call())
        }
    "#,
    &["A"]
);

kotlin_run_test!(
    test_extension_receiver_disambiguates_property,
    r#"
        class Context {
            val label = "root"
            inner class Node {
                val label = "node"
                fun describe(): String = "${this@Context.label}/${label}"
            }
        }

        fun main() {
            println(Context().Node().describe())
        }
    "#,
    &["root/node"]
);

kotlin_run_test!(
    test_apply_with_outer_this_in_block,
    r#"
        class Profile(val base: String) {
            val id: String = "x"
            fun build(): String =
                StringBuilder().apply {
                    this@Profile.base.let { append(it) }
                    append(":")
                    append(id)
                }.toString()
        }

        fun main() {
            println(Profile("p").build())
        }
    "#,
    &["p:x"]
);

kotlin_run_test!(
    test_with_receiver_shadowing,
    r#"
        data class Holder(val value: String)

        fun main() {
            val out = Holder("inner").run {
                val value = "local"
                println(value)
                this.value
            }
            println(out)
        }
    "#,
    &["local", "inner"]
);

kotlin_run_test!(
    test_this_label_in_extension_lambda,
    r#"
        class Host {
            fun transform(): String {
                val f: Host.() -> String = {
                    "${this::class.simpleName}"
                }
                return f()
            }
        }

        fun main() {
            println(Host().transform())
        }
    "#,
    &["Host"]
);

kotlin_run_test!(
    test_outer_this_used_after_nested_block,
    r#"
        class Gate {
            val id = "gate"
            inner class Guard {
                fun value(): String {
                    return this@Gate.run { id }
                }
            }
        }

        fun main() {
            println(Gate().Guard().value())
        }
    "#,
    &["gate"]
);

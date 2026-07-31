use crate::helpers::run_prints;

const REFLECTION_RUNTIME_SETUP: &str = r#"
class Probe(val name: String)

interface MarkerContract

class ProbeImpl(val id: Int) : MarkerContract {
    override fun toString(): String = "ProbeImpl:" + id
}
"#;

fn run_reflection_runtime(src: &str) -> Vec<String> {
    run_prints(&format!("{REFLECTION_RUNTIME_SETUP}\n{src}"))
}

#[test]
fn test_reflection_class_simple_name() {
    let out = run_reflection_runtime(r#"
        fun main() {
            val value = Probe("a")
            println(value::class.simpleName)
            println(Probe::class.simpleName)
        }
    "#);
    assert_eq!(out, &["Probe", "Probe"]);
}

#[test]
fn test_reflection_is_instance_checks() {
    let out = run_reflection_runtime(r#"
        fun main() {
            val value = ProbeImpl(1)
            println(Probe::class.isInstance(value))
            println(MarkerContract::class.isInstance(value))
            println(Probe::class.isInstance("x"))
            println(ProbeImpl::class.isInstance(value))
        }
    "#);
    assert_eq!(out, &["true", "true", "false", "true"]);
}

#[test]
fn test_reflection_java_class_name() {
    let out = run_reflection_runtime(r#"
        fun main() {
            val cls = ProbeImpl::class.java
            println(cls.name)
            println(cls.canonicalName)
            println(cls.simpleName)
        }
    "#);
    assert_eq!(out, &["languages.kotlin.tests.kotlin.test_kotlin_reflection_runtime.ProbeImpl", "languages.kotlin.tests.kotlin.test_kotlin_reflection_runtime.ProbeImpl", "ProbeImpl"]);
}

#[test]
fn test_reflection_object_instance_class_equality() {
    let out = run_reflection_runtime(r#"
        fun main() {
            val a: Any = Probe("x")
            val b: Any = Probe("y")
            println(a::class == b::class)
            println(a::class == Probe::class)
            println(a::class.java == b::class.java)
        }
    "#);
    assert_eq!(out, &["true", "true", "true"]);
}

#[test]
fn test_reflection_array_and_primitive_class_names() {
    let out = run_reflection_runtime(r#"
        fun main() {
            println(Int::class.simpleName)
            println(IntArray::class.simpleName)
            println(Array<Int>::class.simpleName)
            println(String::class.qualifiedName?.endsWith("kotlin.String"))
        }
    "#);
    assert_eq!(out, &["Int", "IntArray", "Array", "true"]);
}

#[test]
fn test_reflection_qualified_name_vs_simple_name() {
    let out = run_reflection_runtime(r#"
        fun main() {
            val c = MarkerContract::class
            println(c.qualifiedName?.contains("MarkerContract"))
            println(c.simpleName)
        }
    "#);
    assert_eq!(out, &["true", "MarkerContract"]);
}

#[test]
fn test_reflection_casting_with_class_refs() {
    let out = run_reflection_runtime(r#"
        fun main() {
            val value: Any = ProbeImpl(7)
            val cls = ProbeImpl::class
            val casted = cls.java.cast(value)
            println(casted is ProbeImpl)
            println((casted as ProbeImpl).id)
        }
    "#);
    assert_eq!(out, &["true", "7"]);
}

#[test]
fn test_reflection_generic_type_reference() {
    let out = run_reflection_runtime(r#"
        fun main() {
            val probe = Probe("abc")
            val values: List<Any> = listOf(probe, 1, "x")
            println(values.map { it::class.simpleName }.joinToString(","))
            val first = values[0]::class
            println(first == Probe::class)
        }
    "#);
    assert_eq!(out, &["Probe,Int,String", "true"]);
}

#[test]
fn test_reflection_property_reference_to_kclass() {
    let out = run_reflection_runtime(r#"
        fun main() {
            val ref = Probe::class
            println(ref.isInstance(Probe("id")))
            println(ref.isInstance(123))
            println(ref.toString().contains("KClass"))
        }
    "#);
    assert_eq!(out, &["true", "false", "true"]);
}

#[test]
fn test_reflection_when_as_result_type() {
    let out = run_reflection_runtime(r#"
        fun main() {
            val values: List<Any> = listOf(Probe("x"), ProbeImpl(1), "str")
            var count = 0
            for (value in values) {
                when {
                    Probe::class.isInstance(value) -> count += 1
                    MarkerContract::class.isInstance(value) -> count += 10
                    else -> count += 100
                }
            }
            println(count)
        }
    "#);
    assert_eq!(out, &["11"]);
}

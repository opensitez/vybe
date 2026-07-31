use crate::helpers::run_prints;

const ANNOTATION_RUNTIME_SETUP: &str = r#"
@Retention(AnnotationRetention.RUNTIME)
@Target(AnnotationTarget.CLASS, AnnotationTarget.FUNCTION, AnnotationTarget.FIELD, AnnotationTarget.VALUE_PARAMETER, AnnotationTarget.PROPERTY)
annotation class Marker(val kind: String)

class AnnotatedModel {
    @field:Marker("field")
    val name: String = "alpha"

    @Marker("ctor")
    fun tagged(@Marker("param") code: Int): Int = code + 1
}

@Marker("service")
class Service
"#;

fn run_annotations_runtime(src: &str) -> Vec<String> {
    run_prints(&format!("{ANNOTATION_RUNTIME_SETUP}\n{src}"))
}

#[test]
fn test_runtime_class_annotation_retrieved() {
    let out = run_annotations_runtime(r#"
        fun main() {
            val ann = Service::class.java.getAnnotation(Marker::class.java)
            println(ann?.kind)
            println(ann != null)
        }
    "#);
    assert_eq!(out, &["service", "true"]);
}

#[test]
fn test_runtime_method_annotation_retrieved() {
    let out = run_annotations_runtime(r#"
        fun main() {
            val method = AnnotatedModel::class.java.getDeclaredMethod("tagged", Int::class.java)
            val ann = method.getAnnotation(Marker::class.java)
            println(ann?.kind)
            println(method.name)
        }
    "#);
    assert_eq!(out, &["ctor", "tagged"]);
}

#[test]
fn test_runtime_parameter_annotation_retrieved() {
    let out = run_annotations_runtime(r#"
        fun main() {
            val method = AnnotatedModel::class.java.getDeclaredMethod("tagged", Int::class.java)
            val param = method.parameters[0]
            val ann = param.getAnnotation(Marker::class.java)
            println(ann?.kind)
            println(param.type.name)
        }
    "#);
    assert_eq!(out, &["param", "int"]);
}

#[test]
fn test_runtime_field_annotation_retrieved() {
    let out = run_annotations_runtime(r#"
        fun main() {
            val field = AnnotatedModel::class.java.getDeclaredField("name")
            val ann = field.getAnnotation(Marker::class.java)
            println(ann?.kind)
            println(field.name)
        }
    "#);
    assert_eq!(out, &["field", "name"]);
}

#[test]
fn test_annotations_apply_to_multiple_targets() {
    let out = run_annotations_runtime(r#"
        class User {
            @Marker("value")
            var value = 0

            @Marker("action")
            fun action() {}
        }

        fun main() {
            val field = User::class.java.getDeclaredField("value")
            val method = User::class.java.getDeclaredMethod("action")
            val a1 = field.getAnnotation(Marker::class.java)
            val a2 = method.getAnnotation(Marker::class.java)
            println(a1?.kind)
            println(a2?.kind)
        }
    "#);
    assert_eq!(out, &["value", "action"]);
}

#[test]
fn test_annotation_on_interface_not_present() {
    let out = run_annotations_runtime(r#"
        interface SampleInterface

        class Impl : SampleInterface

        fun main() {
            val ann = SampleInterface::class.java.getAnnotation(Marker::class.java)
            val implAnn = Impl::class.java.getAnnotation(Marker::class.java)
            println(ann == null)
            println(implAnn == null)
        }
    "#);
    assert_eq!(out, &["true", "true"]);
}

#[test]
fn test_annotation_array_inheritance_not_automatic() {
    let out = run_annotations_runtime(r#"
        @Marker("base")
        open class Base

        class Child : Base()

        fun main() {
            val base = Base::class.java.getAnnotation(Marker::class.java)
            val child = Child::class.java.getAnnotation(Marker::class.java)
            println(base?.kind)
            println(child == null)
        }
    "#);
    assert_eq!(out, &["base", "true"]);
}

#[test]
fn test_annotation_retrieval_after_instance_creation() {
    let out = run_annotations_runtime(r#"
        fun main() {
            val service = Service()
            val annClass = service::class.java.getAnnotation(Marker::class.java)
            val annName = service::class.java.name
            println(annClass?.kind)
            println(annName.endsWith("Service"))
        }
    "#);
    assert_eq!(out, &["service", "true"]);
}

#[test]
fn test_annotation_multiple_calls_stable() {
    let out = run_annotations_runtime(r#"
        fun main() {
            val first = Service::class.java.getAnnotation(Marker::class.java)
            val second = Service::class.java.getAnnotation(Marker::class.java)
            println(first == second)
            println(first?.kind)
        }
    "#);
    assert_eq!(out, &["true", "service"]);
}

#[test]
fn test_annotation_with_local_class() {
    let out = run_annotations_runtime(r#"
        fun main() {
            @Marker("local")
            class Local
            val ann = Local::class.java.getAnnotation(Marker::class.java)
            println(ann?.kind)
        }
    "#);
    assert_eq!(out, &["local"]);
}

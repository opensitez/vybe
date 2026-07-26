use super::helpers::run_python;

#[test]
fn test_python_weakref_callback_called() {
    let src = r#"
import gc
import weakref

flags = {'called': False}
class O:
    pass

o = O()

def cb(ref):
    flags['called'] = True

r = weakref.ref(o, cb)
del o
gc.collect()
print(flags['called'])
"#;
    assert_eq!(run_python(src), vec!["True"]);
}

#[test]
fn test_python_weakref_finalize() {
    let src = r#"
import gc
import weakref

class C: pass
flags = {'cleared': False}
obj = C()
weakref.finalize(obj, lambda: flags.__setitem__('cleared', True))
del obj
gc.collect()
print(flags['cleared'])
"#;
    assert_eq!(run_python(src), vec!["True"]);
}

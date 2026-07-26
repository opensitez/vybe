use super::helpers::run_python;

#[test]
fn test_python_descriptor_data_descriptor() {
    let src = r#"
class Double:
    def __set_name__(self, owner, name):
        self.name = name
    def __get__(self, obj, objtype=None):
        return obj.__dict__.get(self.name, 0) * 2
    def __set__(self, obj, value):
        obj.__dict__[self.name] = value

class C:
    value = Double()

c = C()
c.value = 3
print(c.value)
"#;
    assert_eq!(run_python(src), vec!["6"]);
}

#[test]
fn test_python_property_descriptor() {
    let src = r#"
class Box:
    def __init__(self):
        self._x = 1
    @property
    def x(self):
        return self._x
    @x.setter
    def x(self, v):
        self._x = v

b = Box()
b.x = 7
print(b.x)
"#;
    assert_eq!(run_python(src), vec!["7"]);
}

use super::helpers::{compile_ok, parse_ok, run_prints};
macro_rules! c {
    ($n:ident,$s:expr) => {
        #[test]
        fn $n() {
            compile_ok($s);
        }
    };
}
c!(
    allocate_statement_01,
    "program p
integer, allocatable :: a(:)
allocate(a(3))
end program p
"
);
c!(
    allocate_statement_02,
    "program p
real, allocatable :: a(:,:)
allocate(a(2,2))
end program p
"
);
c!(
    allocate_statement_03,
    "program p
character(len=:), allocatable :: s
allocate(character(len=3) :: s)
end program p
"
);
c!(
    allocate_statement_04,
    "program p
integer, pointer :: p(:)
allocate(p(3))
end program p
"
);
c!(
    allocate_statement_05,
    "program p
integer, allocatable :: x
allocate(x)
end program p
"
);
c!(
    allocate_statement_06,
    "program p
complex, allocatable :: z(:)
allocate(z(2))
end program p
"
);
c!(
    allocate_statement_07,
    "program p
logical, allocatable :: l(:)
allocate(l(2))
end program p
"
);
c!(
    allocate_statement_08,
    "program p
integer, allocatable :: a(:,:,:)
allocate(a(2,2,2))
end program p
"
);
c!(
    allocate_statement_09,
    "program p
type t
 integer :: x
end type t
type(t), allocatable :: a(:)
allocate(a(2))
end program p
"
);
c!(
    allocate_statement_10,
    "program p
class(*), allocatable :: x
allocate(integer :: x)
end program p
"
);

#[test]
fn allocate_statement_runtime_array_fill_and_sum() {
    assert_eq!(
        run_prints(
            "program t\n\
integer, allocatable :: a(:)\n\
allocate(a(3))\n\
a = [10, 20, 30]\n\
print *, sum(a)\n\
end program t\n"
        ),
        vec!["60"]
    );
}

#[test]
fn allocate_statement_runtime_pointer_array_and_modify() {
    assert_eq!(
        run_prints(
            "program t\n\
integer, pointer :: p(:)\n\
allocate(p(2))\np(1) = 7\n\
p(2) = 9\n\
print *, p(2)\n\
deallocate(p)\n\
print *, 'done'\n\
end program t\n"
        ),
        vec!["9", "done"]
    );
}

#[test]
fn allocate_statement_runtime_character_alloc_and_length() {
    assert_eq!(
        run_prints(
            "program t\n\
character(len=:), allocatable :: s\n\
allocate(character(len=4) :: s)\n\
s = 'fort'\n\
print *, len(s)\n\
print *, s\n\
end program t\n"
        ),
        vec!["4", "fort"]
    );
}

#[test]
fn allocate_statement_parse_rejects_missing_array_shape() {
    assert!(!parse_ok(
        "program p
integer, allocatable :: a(:)
allocate(a)
end program p
"
    ));
}

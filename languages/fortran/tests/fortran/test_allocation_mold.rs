use super::helpers::{compile_ok, run_prints};
macro_rules! c {
    ($n:ident,$s:expr) => {
        #[test]
        fn $n() {
            compile_ok($s);
        }
    };
}
c!(
    allocation_mold_01,
    "program p
integer, allocatable :: a(:), b(:)
allocate(b(2))
allocate(a, mold=b)
end program p
"
);
c!(
    allocation_mold_02,
    "program p
integer, allocatable :: a,b
allocate(b)
allocate(a, mold=b)
end program p
"
);
c!(
    allocation_mold_03,
    "program p
real, allocatable :: a(:,:), b(:,:)
allocate(b(2,2))
allocate(a, mold=b)
end program p
"
);
c!(
    allocation_mold_04,
    "program p
character(len=:), allocatable :: s, t
allocate(character(len=4) :: t)
allocate(s, mold=t)
end program p
"
);
c!(
    allocation_mold_05,
    "program p
logical, allocatable :: a(:), b(:)
allocate(b(2))
allocate(a, mold=b)
end program p
"
);
c!(
    allocation_mold_06,
    "program p
complex, allocatable :: a(:), b(:)
allocate(b(2))
allocate(a, mold=b)
end program p
"
);
c!(
    allocation_mold_07,
    "type t
 integer :: x
end type t
program p
type(t), allocatable :: a,b
allocate(b)
allocate(a, mold=b)
end program p
"
);
c!(
    allocation_mold_08,
    "program p
class(*), allocatable :: x, y
allocate(integer :: y)
allocate(x, mold=y)
end program p
"
);
c!(
    allocation_mold_09,
    "program p
integer, pointer :: p(:)
integer, allocatable :: a(:)
allocate(p(3))
allocate(a, mold=p)
end program p
"
);
c!(
    allocation_mold_10,
    "program p
real, allocatable :: a(:,:,:), b(:,:,:)
allocate(b(2,2,2))
allocate(a, mold=b)
end program p
"
);

#[test]
fn allocation_mold_copies_array_payload_and_shape() {
    assert_eq!(
        run_prints(
            "program t\n\
integer, allocatable :: a(:), b(:)\n\
allocate(b(3))\n\
b = [2, 4, 6]\n\
allocate(a, mold=b)\n\
print *, size(a)\n\
print *, a(1)\n\
print *, a(2)\n\
print *, a(3)\n\
end program t\n"
        ),
        vec!["3", "2", "4", "6"]
    );
}

#[test]
fn allocation_mold_copies_scalar_payload() {
    assert_eq!(
        run_prints(
            "program t\n\
integer, allocatable :: a, b\n\
allocate(b)\n\
b = 17\n\
allocate(a, mold=b)\n\
print *, a\n\
end program t\n"
        ),
        vec!["17"]
    );
}

#[test]
fn allocation_mold_copies_character_payload() {
    assert_eq!(
        run_prints(
            "program t\n\
character(len=:), allocatable :: a, b\n\
allocate(character(len=5) :: b)\n\
b = 'moldx'\n\
allocate(a, mold=b)\n\
print *, len(a)\n\
print *, a\n\
end program t\n"
        ),
        vec!["5", "moldx"]
    );
}

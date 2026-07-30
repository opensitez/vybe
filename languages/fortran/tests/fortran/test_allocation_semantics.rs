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
    as_01,
    "program p
integer, allocatable :: a(:)
allocate(a(3))
print *, size(a)
end program p
"
);
c!(
    as_02,
    "program p
integer, allocatable :: a(:)
allocate(a(3), source=[1,2,3])
print *, a
end program p
"
);
c!(
    as_03,
    "program p
integer, allocatable :: a(:),b(:)
allocate(b(2))
allocate(a, mold=b)
end program p
"
);
c!(
    as_04,
    "program p
integer, allocatable :: a(:),b(:)
allocate(a(3))
call move_alloc(a,b)
end program p
"
);
c!(
    as_05,
    "program p
integer, pointer :: p(:)
allocate(p(3))
end program p
"
);
c!(
    as_06,
    "program p
real, allocatable :: a(:,:)
allocate(a(2,3))
end program p
"
);
c!(
    as_07,
    "program p
complex, allocatable :: a(:)
allocate(a(4))
end program p
"
);
c!(
    as_08,
    "program p
logical, allocatable :: a(:)
allocate(a(5))
end program p
"
);
c!(
    as_09,
    "program p
character(len=:), allocatable :: s
allocate(character(len=5) :: s)
end program p
"
);
c!(
    as_10,
    "program p
integer, allocatable :: x
allocate(x, source=5)
end program p
"
);
c!(
    as_11,
    "program p
integer, allocatable :: a(:)
integer :: st
allocate(a(3), stat=st)
end program p
"
);
c!(
    as_12,
    "program p
integer, allocatable :: a(:)
integer :: st
allocate(a(3))
deallocate(a, stat=st)
end program p
"
);
c!(
    as_13,
    "type :: t
integer, allocatable :: a(:)
end type t
program p
type(t) :: x
allocate(x%a(2))
end program p
"
);
c!(
    as_14,
    "type :: t
character(len=:), allocatable :: s
end type t
program p
type(t) :: x
allocate(character(len=3) :: x%s)
end program p
"
);
c!(
    as_15,
    "program p
integer, allocatable :: a(:,:,:)
allocate(a(2,2,2))
end program p
"
);
c!(
    as_16,
    "program p
integer, allocatable :: a(:,:,:,:)
allocate(a(2,2,2,2))
end program p
"
);
c!(
    as_17,
    "program p
integer, allocatable :: a(:)
allocate(a(0))
print *, size(a)
end program p
"
);
c!(
    as_18,
    "program p
integer, allocatable :: a(:)
allocate(a(3))
deallocate(a)
end program p
"
);
c!(
    as_19,
    "program p
integer, allocatable :: a(:)
allocate(a(3))
a = [1,2,3]
end program p
"
);
c!(
    as_20,
    "program p
class(*), allocatable :: x
allocate(integer :: x)
end program p
"
);

#[test]
fn allocation_semantics_runtime_move_alloc_transfers_allocation() {
    assert_eq!(
        run_prints(
            "program t\n\
integer, allocatable :: src(:), dst(:)\n\
allocate(src(2))\nsrc = [8, 9]\n\
call move_alloc(src, dst)\n\
print *, dst(1)\n\
print *, dst(2)\n\
print *, allocated(src)\n\
print *, allocated(dst)\n\
end program t\n"
        ),
        vec!["8", "9", "false", "true"]
    );
}

#[test]
fn allocation_semantics_runtime_zero_extent_returns_size_zero() {
    assert_eq!(
        run_prints(
            "program t\n\
integer, allocatable :: a(:)\n\
allocate(a(0))\n\
print *, size(a)\n\
deallocate(a)\n\
end program t\n"
        ),
        vec!["0"]
    );
}

#[test]
fn allocation_semantics_runtime_allocates_derived_component_field() {
    assert_eq!(
        run_prints(
            "type :: holder\n\
character(len=:), allocatable :: s\n\
end type holder\n\
program t\n\
type(holder) :: h\n\
allocate(character(len=5) :: h%s)\n\
h%s = 'abcde'\n\
print *, len(h%s)\n\
print *, h%s\n\
end program t\n"
        ),
        vec!["5", "abcde"]
    );
}

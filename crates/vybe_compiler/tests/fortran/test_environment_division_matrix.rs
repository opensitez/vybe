use super::helpers::compile_ok;

macro_rules! c {
    ($name:ident, $src:expr) => {
        #[test]
        fn $name() { compile_ok($src); }
    };
}

c!(env_implicit_none_01, "program p
implicit none
integer :: x
x = 1
print *, x
end program p
");
c!(env_implicit_real_02, "program p
implicit real(a-h,o-z)
a = 1.0
print *, a
end program p
");
c!(env_use_assoc_03, "module m
integer :: x = 1
end module m
program p
use m
print *, x
end program p
");
c!(env_use_only_04, "module m
integer :: x = 1, y = 2
end module m
program p
use m, only: x
print *, x
end program p
");
c!(env_use_rename_05, "module m
integer :: x = 1
end module m
program p
use m, y => x
print *, y
end program p
");
c!(env_use_intrinsic_06, "program p
use, intrinsic :: iso_fortran_env
print *, int8
end program p
");
c!(env_use_nonintrinsic_07, "module m
integer :: x = 1
end module m
program p
use, non_intrinsic :: m
print *, x
end program p
");
c!(env_host_assoc_08, "program p
integer :: x
x = 1
call s()
contains
subroutine s()
print *, x
end subroutine s
end program p
");
c!(env_import_09, "module m
integer :: x
contains
subroutine s()
 import :: x
 integer :: y
 y = x
end subroutine s
end module m
");
c!(env_spec_order_10, "program p
implicit none
integer, parameter :: n = 3
integer :: a(n)
print *, 1
end program p
");
c!(env_forward_ref_11, "program p
implicit none
integer :: a(n)
integer, parameter :: n = 3
print *, 1
end program p
");
c!(env_namelist_12, "program p
implicit none
integer :: x
namelist /grp/ x
print *, 1
end program p
");
c!(env_common_13, "program p
implicit none
integer :: x
common /blk/ x
print *, 1
end program p
");
c!(env_equivalence_14, "program p
implicit none
integer :: a, b
equivalence(a,b)
print *, 1
end program p
");
c!(env_block_data_15, "block data bd
implicit none
integer :: x
common /blk/ x
data x /1/
end block data bd
");
c!(env_processor_dep_16, "program p
implicit none
integer :: x
print *, x
end program p
");
c!(env_diag_17, "program p
implicit none
integer :: x
x = 1
print *, x
end program p
");
c!(env_private_public_18, "module m
implicit none
private
public :: x
integer :: x
end module m
");
c!(env_scope_module_19, "module m
implicit none
integer :: x
contains
subroutine s()
print *, x
end subroutine s
end module m
");
c!(env_scope_block_20, "program p
block
integer :: x
x = 1
print *, x
end block
end program p
");
c!(env_scope_assoc_21, "program p
integer :: x
x = 1
associate(y => x)
 print *, y
end associate
end program p
");
c!(env_scope_construct_22, "program p
integer :: i
do i = 1, 1
 print *, i
end do
end program p
");
c!(env_exec_sequence_23, "program p
integer :: x
x = 1
x = x + 1
print *, x
end program p
");
c!(env_storage_assoc_24, "program p
integer :: a(2)
integer :: b
equivalence(a(1), b)
print *, 1
end program p
");
c!(env_conformance_25, "program p
implicit none
integer :: x
x = 1
print *, x
end program p
");
c!(env_constraint_26, "program p
implicit none
integer :: x
x = 1
print *, x
end program p
");
c!(env_syntax_27, "program p
implicit none
print *, 1
end program p
");
c!(env_semantic_28, "program p
implicit none
integer :: x
x = 1
print *, x
end program p
");
c!(env_runtime_29, "program p
implicit none
integer :: x
x = 1
print *, x
end program p
");
c!(env_opt_barrier_30, "program p
implicit none
volatile :: x
integer :: x
x = 1
print *, x
end program p
");
! vybe-test: fortran/reduce_intrinsic/reduce_custom_min_procedure
! origin: languages/fortran/tests/fortran/test_fortran2018_extended.rs
! The REDUCE call is hoisted out of the IF condition: gfortran 16.1 ICEs
! (internal compiler error in gfc_typenode_for_spec, trans-types.cc:1331)
! when REDUCE appears directly inside one. The Fortran was already correct —
! a PURE function of two arguments is exactly what F2018 REDUCE requires —
! so this is a workaround for the ground-truth compiler, not a fix to the test.
program t
integer :: a(5) = [8, 3, 9, 1, 6]
vy_r = reduce(a, pick_min)
if ((vy_r) /= 1) then
    print *, "FAIL: want [1] got [", vy_r, "]"
    stop 1
end if
contains
pure function pick_min(x, y) result(r)
integer, intent(in) :: x, y
integer :: r
integer :: vy_r
r = min(x, y)
end function pick_min
end program t

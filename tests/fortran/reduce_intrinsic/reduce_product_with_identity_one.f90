! vybe-test: fortran/reduce_intrinsic/reduce_product_with_identity_one
! origin: languages/fortran/tests/fortran/test_fortran2018_extended.rs
! `operator(+)` is a GENERIC-SPEC — legal in INTERFACE/USE/generic bindings,
! NOT as an actual argument. F2018 REDUCE takes a PURE FUNCTION of two args.
! The result is also hoisted into a variable: gfortran 16.1 ICEs
! (gfc_typenode_for_spec, trans-types.cc:1331) when REDUCE appears directly
! inside an IF condition. Same value, and it compiles.
program t
integer :: a(4) = [2, 3, 4, 5]
integer :: vy_r
vy_r = reduce(a, vy_mul, identity=1)
if ((vy_r) /= 120) then
    print *, "FAIL: want [120] got [", vy_r, "]"
    stop 1
end if
contains
    pure function vy_mul(x, y) result(r)
        integer, intent(in) :: x, y
        integer :: r
        r = x * y
    end function vy_mul
end program t

! vybe-test: fortran/reduce_intrinsic/reduce_dim1_with_mask_columns
! origin: languages/fortran/tests/fortran/test_fortran2018_extended.rs
! `operator(+)` is a GENERIC-SPEC — legal in INTERFACE/USE/generic bindings,
! NOT as an actual argument. F2018 REDUCE takes a PURE FUNCTION of two args.
! The result is also hoisted into a variable: gfortran 16.1 ICEs
! (gfc_typenode_for_spec, trans-types.cc:1331) when REDUCE appears directly
! inside an IF condition. Same value, and it compiles.
program t
integer :: m(2,3) = reshape([1,2,3,4,5,6],[2,3])
logical :: mask(2,3) = reshape([.true.,.false.,.true.,.false.,.true.,.false.],[2,3])
integer :: r(3)
r = reduce(m, vy_add, dim=1, mask=mask)
if ((r(1)) /= 1) then
    print *, "FAIL: want [1] got [", r(1), "]"
    stop 1
end if
if ((r(2)) /= 3) then
    print *, "FAIL: want [3] got [", r(2), "]"
    stop 1
end if
if ((r(3)) /= 5) then
    print *, "FAIL: want [5] got [", r(3), "]"
    stop 1
end if
contains
    pure function vy_add(x, y) result(r)
        integer, intent(in) :: x, y
        integer :: r
        r = x + y
    end function vy_add
end program t

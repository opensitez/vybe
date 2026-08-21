! vybe-test: fortran/reduce_intrinsic/reduce_logical_and_chain
! origin: languages/fortran/tests/fortran/test_fortran2018_extended.rs
! `operator(+)` is a GENERIC-SPEC — legal in INTERFACE/USE/generic bindings,
! NOT as an actual argument. F2018 REDUCE takes a PURE FUNCTION of two args.
! The result is also hoisted into a variable: gfortran 16.1 ICEs
! (gfc_typenode_for_spec, trans-types.cc:1331) when REDUCE appears directly
! inside an IF condition. Same value, and it compiles.
program t
logical :: flags(3) = [.true., .true., .false.]
logical :: vy_r
vy_r = reduce(flags, vy_and)
if ((vy_r) .neqv. .false.) then
    print *, "FAIL: want [false] got [", vy_r, "]"
    stop 1
end if
contains
    pure function vy_and(x, y) result(r)
        logical, intent(in) :: x, y
        logical :: r
        r = x .and. y
    end function vy_and
end program t

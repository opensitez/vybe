! vybe-test: fortran/fortran2018_extended/reduce_logical_and_chain
! origin: languages/fortran/tests/fortran/test_fortran2018_extended.rs
program t
logical :: flags(3) = [.true., .true., .false.]
if ((reduce(flags, operator(.and.))) .neqv. .false.) then
    print *, "FAIL: want [false] got [", reduce(flags, operator(.and.)), "]"
    stop 1
end if
end program t

! vybe-test: fortran/fortran2003_extended/alloc_comp_logical_flag_value
! origin: languages/fortran/tests/fortran/test_fortran2003_extended.rs
program t
type :: Flags
logical, allocatable :: bits(:)
end type Flags
type(Flags) :: f
f%bits = [.true., .false., .true.]
if ((count(f%bits)) /= 2) then
    print *, "FAIL: want [2] got [", count(f%bits), "]"
    stop 1
end if
end program t

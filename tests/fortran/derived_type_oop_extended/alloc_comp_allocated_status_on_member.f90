! vybe-test: fortran/derived_type_oop_extended/alloc_comp_allocated_status_on_member
! origin: languages/fortran/tests/fortran/test_derived_type_oop_extended.rs
program t
type :: Holder
real, allocatable :: vals(:)
end type Holder
type(Holder) :: h
if ((allocated(h%vals)) .neqv. .false.) then
    print *, "FAIL: want [false] got [", allocated(h%vals), "]"
    stop 1
end if
allocate(h%vals(1))
h%vals(1) = 2.5
if ((int(h%vals(1))) /= 2) then
    print *, "FAIL: want [2] got [", int(h%vals(1)), "]"
    stop 1
end if
if ((allocated(h%vals)) .neqv. .true.) then
    print *, "FAIL: want [true] got [", allocated(h%vals), "]"
    stop 1
end if
end program t

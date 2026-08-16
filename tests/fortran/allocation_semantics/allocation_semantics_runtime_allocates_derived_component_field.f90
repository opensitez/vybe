! vybe-test: fortran/allocation_semantics/allocation_semantics_runtime_allocates_derived_component_field
! origin: languages/fortran/tests/fortran/test_allocation_semantics.rs
program t
type :: holder
character(len=:), allocatable :: s
end type holder
type(holder) :: h
allocate(character(len=5) :: h%s)
h%s = 'abcde'
if ((len(h%s)) /= 5) then
    print *, "FAIL: want [5] got [", len(h%s), "]"
    stop 1
end if
if (trim(h%s) /= "abcde") then
    print *, "FAIL: want [abcde] got [", h%s, "]"
    stop 1
end if
end program t

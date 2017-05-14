#resolver no JuMP

include("refine.jl")
mod = Model(solver=ClpSolver()) #tipo do modelo
@variable(mod,x[1:61]>=0)
@constraint(mod,A*x.<=u)
@objective(mod,Min,dot(f,x))

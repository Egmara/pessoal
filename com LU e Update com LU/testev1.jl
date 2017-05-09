using Clp, JuMP, MathProgBase

function testev1(fun)
mod = Model(solver=ClpSolver())
m_internal = MathProgBase.LinearQuadraticModel(ClpSolver()) #tipo do solver

include("funcoes2.jl")
fim = length(funcoes2)
#for fun=1:length(funcoes) #para cada funçao
#for fun=23:23

    MathProgBase.loadproblem!(m_internal, funcoes2[fun]) #carrega os dados da função

    f = MathProgBase.getobj(m_internal)
    A = MathProgBase.getconstrmatrix(m_internal)
    m, n = size(A)
    xlb = MathProgBase.getvarLB(m_internal)
    xub = MathProgBase.getvarUB(m_internal)
    l = MathProgBase.getconstrLB(m_internal)
    u = MathProgBase.getconstrUB(m_internal)
    b = zeros(m)

    cont = 0

    pode = true;

    println(funcoes2[fun])
    println("m = $m e n = $n e fun = $fun")
    println(size(A))

  for i = 1:m
      c = zeros(m)
    	#println(" l:$(l[i]); u:$(u[i])")
      if l[i] == -Inf
      		if u[i] == Inf
        		pode = false;
        	  break;
      		elseif u[i] == -Inf
        		pode = false;
        	  break;
          else
            c[i] = 1.0
        	  b[i] = u[i]
            A = [A c]
        	  #u[i]!=+-Inf
        	  #println("<=")
            cont = cont + 1
          end
      elseif l[i] == Inf
      		pode = false;
          break;
      else #l[i]!=+-Inf
      		if u[i] != Inf
        		if u[i] == l[i] #u!=Inf
              b[i] = l[i]
       			else
          		pode = false; #(l!=Inf & u!=Inf) & u!=l
          		break;
        		end
      		 else
       		    b[i] = l[i]
              c[i] = -1.0
              A = [A c]
        	     #u=Inf
        	     #println(">=")
                cont = cont + 1
        	 end
      end
        #println(size(A))
    end

    f = [f;zeros(cont)]
    #println(b)
    A = sparse(full(A))

    #ClpSolver
    #@variable(mod,x[1:n+cont]>=0)
    #@constraint(mod,A*x.==b)
    #@objective(mod,Min,dot(f,x))
    #solve(mod)
    #xv = getvalue(x)
    #println("objetivo = $(dot(f,xv))")

    #teste
    #include("Simplex_LUfact.jl")
    #@time x, base, nbase, fx = SSolveFact(f,A,b)
    #println("x = $x")
    #println("fx com LU fact = $fx")

    include("Simplex_Rev_comLU.jl")
    @time x1, base1, nbase1, fx1 = SSolve(f,A,b)

    #include("Simplex_Rev_UpdateFact.jl")
    #@time x2, base2, nbase2, fx2 = SSolveUpFact(f,A,b)
    #println("x2 = $x2")
    #println("fx com Update fact = $fx2")

    include("Simplex_Rev_Update.jl")
    @time x3, base3, nbase3, fx3 = SSolveUp(f,A,b)
    println("x3 = $x3")
    println("fx com Update = $fx3")

    println(funcoes2[fun])
    println("m = $m e n = $n e fun = $fun")
    println("cont = $cont")
    println(size(A))

  end

using Clp, JuMP, MathProgBase

mod = Model(solver=ClpSolver())
m_internal = MathProgBase.LinearQuadraticModel(ClpSolver()) #tipo do solver

function testev4(opt,funcao)
#include("funcoes2.jl")
#fim = length(funcoes2)
#for fun=1:length(funcoes) #para cada funçao
#for fun=36:36

    MathProgBase.loadproblem!(m_internal, funcao) #carrega os dados da função

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

    #println(funcao)
    println("m = $m e n = $n e fun = $funcao")
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
    if opt==1
    include("Simplex_Rev_comLU.jl")
    tempo=@elapsed x, base, nbase, fx = SSolve(f,A,b)
    end
    if opt==2
    include("Simplex_Rev_Update.jl")
    tempo=@elapsed x, base, nbase, fx = SSolveUp(f,A,b)
    end
    if opt==3
    include("Simplex_Rev_UpdateFact.jl")
    tempo=@elapsed x, base, nbase, fx = SSolveUpFact(f,A,b)
  end
    if opt==4
    include("Simplex_LUfact.jl")
    tempo=@elapsed x, base, nbase, fx = SSolveFact(f,A,b)
  end

    #println(funcao)
    #println("m = $m e n = $n e fun = $funcao")
    #println("cont = $cont")
    #println(size(A))
    return fx,tempo,m,n,cont
  end

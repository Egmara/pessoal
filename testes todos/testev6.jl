using Clp, JuMP, MathProgBase

mod = Model(solver=ClpSolver())
m_internal = MathProgBase.LinearQuadraticModel(ClpSolver()) #tipo do solver

#lista=["adlittle","share2b","scagr7","p0282","p0548 ","p0291","agg2","agg3","nsic1"]
#include("funcoes2.jl")
#fim = length(funcoes2)
#for fun=1:length(funcoes) #para cada funçao
function testev6_1()
  lista=["kleemin3.mps","kleemin4.mps","kleemin5.mps","kleemin6.mps","kleemin7.mps","kleemin8.mps","farm.mps","afiro.mps","refine.mps","sc50a.mps","sc50b.mps"]
  Fxi=zeros(length(lista),4)
  Tempoi=zeros(length(lista),4)
  Dadosi=zeros(length(lista),3)
for fun=1:length(lista)
  funcao=lista[fun]

    MathProgBase.loadproblem!(m_internal, funcao); #carrega os dados da função

    f = MathProgBase.getobj(m_internal);
    A = MathProgBase.getconstrmatrix(m_internal);
    m, n = size(A);
    xlb = MathProgBase.getvarLB(m_internal);
    xub = MathProgBase.getvarUB(m_internal);
    l = MathProgBase.getconstrLB(m_internal);
    u = MathProgBase.getconstrUB(m_internal);
    b = zeros(m);

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
    # 4 tipos de resolução


    include("Simplex_Rev_comLU.jl")
    Tempoi[fun,1] = @elapsed x, base, nbase, Fxi[fun,1] = SSolve(f,A,b)
    #Tempoi[fun,1] = -1.0
    #Fxi[fun,1] = -1.0

    include("Simplex_Rev_Update.jl")
    Tempoi[fun,2] = @elapsed x, base, nbase, Fxi[fun,2] = SSolveUp(f,A,b)


    include("Simplex_Rev_UpdateFact.jl")
    Tempoi[fun,3] = @elapsed x, base, nbase, Fxi[fun,3] = SSolveUpFact(f,A,b)


    include("Simplex_LUfact.jl")
    Tempoi[fun,4] = @elapsed x, base, nbase, Fxi[fun,4] = SSolveFact(f,A,b)

    Dadosi[fun,:] = [m n cont ]

    #println(funcao)
    #println("m = $m e n = $n e fun = $funcao")
    #println("cont = $cont")
    #println(size(A))
    #return Fxi,Tempoi,[m n cont]
  end
return lista,Fxi,Tempoi,Dadosi
end

function testev6_2()
  lista=["share2b.mps","scagr7.mps","p0282.mps","p0548.mps","p0291.mps","agg2.mps","agg3.mps","nsic1.mps"]
  Fxi=zeros(length(lista),4)
  Tempoi=zeros(length(lista),4)
  Dadosi=zeros(length(lista),3)
for fun=1:length(lista)
  funcao=lista[fun]

    MathProgBase.loadproblem!(m_internal, funcao); #carrega os dados da função

    f = MathProgBase.getobj(m_internal);
    A = MathProgBase.getconstrmatrix(m_internal);
    m, n = size(A);
    xlb = MathProgBase.getvarLB(m_internal);
    xub = MathProgBase.getvarUB(m_internal);
    l = MathProgBase.getconstrLB(m_internal);
    u = MathProgBase.getconstrUB(m_internal);
    b = zeros(m);

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
    # 4 tipos de resolução



    include("Simplex_Rev_UpdateFact.jl")
    Tempoi[fun,3] = @elapsed x, base, nbase, Fxi[fun,3] = SSolveUpFact(f,A,b)


    #include("Simplex_LUfact.jl")
    #Tempoi[fun,4] = @elapsed x, base, nbase, Fxi[fun,4] = SSolveFact(f,A,b)

    Dadosi[fun,:] = [m n cont ]

    #println(funcao)
    #println("m = $m e n = $n e fun = $funcao")
    #println("cont = $cont")
    #println(size(A))
    #return Fxi,Tempoi,[m n cont]
  end
return lista,Fxi,Tempoi,Dadosi
end

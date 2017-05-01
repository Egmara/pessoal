using Clp, JuMP, MathProgBase

mod=Model(solver=ClpSolver()) #tipo do modelo
m_internal = MathProgBase.LinearQuadraticModel(ClpSolver()) #tipo do solver
include("testfun.jl")

for fun=1:length(funcoes) #para cada funçao
    MathProgBase.loadproblem!(m_internal, funcoes[fun]) #carrega os dados da função

    xlb = MathProgBase.getvarLB(m_internal)
    xub = MathProgBase.getvarUB(m_internal)
    l = MathProgBase.getconstrLB(m_internal)
    u = MathProgBase.getconstrUB(m_internal)


    pode=true;
    for i=1:length(xlb)
    #println("xlb:$(xlb[i]); xub:$(xub[i])")
      if xlb[i]!=0 || xub[i]!=Inf

        pode=false;
        break;
      end
    end

  for i=1:length(l)
    #println(" l:$(l[i]); u:$(u[i])")
    if l[i]==-Inf
      if u[i]==Inf
        pode=false;
        break;
      elseif u[i]==-Inf
        pode=false;
        break;
      else #u[i]!=+-Inf
        #println("<=")
      end
    elseif l[i]==Inf
      pode=false;
      break;

    else #l[i]!=+-Inf
      if u[i]!=Inf
        if u[i]==l[i] #u!=Inf
          #println("=")

        else
          pode=false; #(l!=Inf & u!=Inf) & u!=l
          break;
        end
      else #u=Inf
        #println(">=")
      end
    end

  end

  if pode
    println(funcoes[fun])
  end

end

include("solve_tri_LU.jl")
include("exercicio2.jl")

function SSolve(c,A,b)

  (m,n) = size(A)

  # Fase 1
  ca = [zeros(n);ones(m)]
  B = [A speye(m)]
  xb = b

  for i = 1:m
    if b[i] < 0
      B[i,n+i] = -1
      xb[i] = -b[i]
    end
  end

  # resolve o Simplex com variáveis artificiais para achar base inicial
  x, base = Simplex_Rev(ca,B,b,collect(n+1:n+m),collect(1:n),xb)
  nbase = zeros(Int64,n-m)

  j = 1
  for i = 1:n
    if all(i.!= base)
      nbase[j] = i
      j = j+1
    end
  end

  # resolve o PL original
  x, base, nbase, fx = Simplex_Rev(c,A,b,base,nbase,x[base])
  return x, base, nbase , fx
end


function Simplex_Rev(c,A,b,base,nbase,xb)

  (m,n) = size(A)
  dimb = length(base)
  dimnb = length(nbase)
  x = zeros(n)
  lambda = zeros(m)
  d = zeros(m)

  # encontra primeira fatoração LU com pivotemento
  L,U,P,Q = exercicio2(A[:,base])
  println(base)
  x[base] = xb
  println("base inicial = $base")
  println("xinicial = $x")

  # Resolver o sistema B'v=cb para achar  o vetor v (lambda tamanho m)
  cb = c[base]
  y = solve_tri_inf(U',cb[Q])
  lambda[P] = solve_tri_sup(L', y)

  # Calcular sn = cn-N'*v (tamanho m - n)
  sn = c[nbase] - A[:,nbase]'*lambda

  #teste
  println("lambda inicial = $lambda")
  println("sn inicial = $sn")

  #i = 0
  #while minimum(sn) < 0
  while minimum(sn) < 0.0 && abs(minimum(sn)) > 1e-9
    #achar os índices q e p
    q = 0
    for k = 1:dimnb
      if sign(sn[k]) == -1 && abs(sn[k]) > 1e-9
  	    q = k
  	    break
      end
    end

    y = solve_tri_inf(L,A[P,nbase[q]])
    d[Q] = solve_tri_sup(U, y)
    println("d=$d")

    if  all(d.<= 0)
     error("O problema é ilimitado!")
    end
    alpha = Inf

    p = 0
    for j = 1:dimb
      if d[j] > 1e-10 && x[base[j]] / d[j] < alpha
        alpha = x[base[j]]/d[j]
        p = j
      end
    end
    println("alpha = $alpha")

    # atualizando o x
    x[base] = x[base] - alpha*d
    # atualizando o x que entra na base
    x[nbase[q]] = alpha
    # depois troca as bases
    nbase[q], base[p] = base[p], nbase[q]
    println("  nbase=$(nbase[q]), base=$(base[p])")
    # encontra fatoração LU
    L,U,P,Q = exercicio2(A[:,base])

    cb = c[base]
    y = solve_tri_inf(U',cb[Q])
    lambda[P] = solve_tri_sup(L', y)
    println("z = $(dot(c,x))")

    sn = c[nbase] - A[:,nbase]'*lambda
	  println("base = $base")
    println("lambda = $lambda")
    println("sn = $sn")

  end

  println("Ponto ótimo encontrado!")
  println("z = $(dot(c,x))")
  return x, base, nbase, dot(c,x) #retornar ponto otimo
end

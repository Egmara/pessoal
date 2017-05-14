include("LUupdate.jl")
include("solve_tri_LU.jl")
include("exercicio2.jl")

function SSolveUp(c,A,b)

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
  x, base = Simplex_Rev_comUp(ca,B,b,collect(n+1:n+m),collect(1:n),xb)
  nbase = zeros(Int64,n-m)

  j = 1
  for i = 1:n
    if all(i.!= base)
      nbase[j] = i
      j = j+1
    end
  end

  # resolve o PL original
  x, base, nbase = Simplex_Rev_comUp(c,A,b,base,nbase,x[base])
  return x, base, nbase, dot(c,x)
end

function Simplex_Rev_comUp(c,A,b,base,nbase,xb)

  (m,n) = size(A)
  dimb = length(base)
  dimnb = length(nbase)
  x = zeros(n)
  lambda = zeros(m)
  d = zeros(m)

  # encontra primeira fatoração LU com pivotiamento
  L0,U,P0,Q0 = exercicio2(A[:,base])

  x[base] = xb
  #println(x)

  # teste
  #println("base inicial = $base")
  #println("xinicial = $x")

  # Resolve o sistema B'lambda = cb (tamanho m)
  cb = c[base]
  z = solve_tri_inf(U',cb[Q0])
  lambda[P0] = solve_tri_sup(L0', z)
  sn = c[nbase] - A[:,nbase]'*lambda

  #teste
  #println("lambda inicial = $lambda")
  #println("sn inicial = $sn")

  # matriz que guarda as permutações
  P = []
  # matriz que guarda os valores da Ltiu inversa
  L = []
  # matriz que guarda r e p
  rp = []

  # numero de updates
  iter = 0

  while minimum(sn) < 0.0 && abs(minimum(sn)) > 1e-10

    q = 0
    for k = 1:dimnb
      #if sn[k] < 0.0
      if sign(sn[k]) == -1 && abs(sn[k]) > 1e-10
        #println(sn[k])
  	    q = k
  	    break
      end
    end

    if iter == 0
      y = solve_tri_inf(L0,A[P0,nbase[q]])
      d[Q0] = solve_tri_sup(U,y)
	    # permutar a base antes de escolher o p
	    base = base[Q0]
	    d = d[Q0]
    else
      y = solve_tri_inf(L0,A[P0,nbase[q]])
      # resolve sistemas acumulados
      for i = 1:iter
        y = y[P[i]]
		    r = rp[i][1]
        y[r] = dot(L[i],y)
      end
      # encontra d
      d = solve_tri_sup(U,y)
    end
    #println("d = $d")

    if  all(d.<= 0)
     error("O problema é ilimitado!")
    end
    alpha = Inf

    p = 0
    for j = 1:dimb
      #1e-10
      if d[j] > 1e-10 && x[base[j]] / d[j] < alpha
        alpha = x[base[j]]/d[j]
        p = j
      end
    end

    # atualizando o x
    x[base] = x[base] - alpha*d
    # atualizando o x que entra na base
    x[nbase[q]] = alpha
    # depois troca as bases
    nbase[q], base[p] = base[p], nbase[q]

    #println("  nbase=$(nbase[q]), base=$(base[p])")

    # último não zero na coluna do spike
    r = dimb
    for nz = dimb:-1:1
      if abs(y[nz]) > 1e-12
        #println(y[nz])
        r = nz
        break
      end
    end

    #println("p = $p e r = $r")
    # atualiza as matrizes
    U,lk,Pk = LUupdate(U,r,p,m,y)
    # guarda as permutações
    push!(P,Pk)
    # guarda os vetores l da inversa da Ltiu
    push!(L,lk)
	  # guarda r e p
	  push!(rp,[r,p])

    # permuta a base antes de escolher o próximo p
    base = base[Pk]
    #println("permuta a base atual= $base")

    # número de updates
    iter = iter + 1

    # encontra próximo lambda Uk'*Lk'v = cb
    z = solve_tri_inf(U', c[base])
    # resolve sistemas acumulados
    for i = iter:-1:1
      Lt = L[i]
      Pi = P[i]
	    r = rp[i][1]
      p = rp[i][2]
      z[p:r-1] = z[p:r-1] + z[r]*Lt[p:r-1]
      z[Pi] = z
    end
    lambda[P0] = solve_tri_sup(L0',z)

    # encontra proximo s
    sn = c[nbase] - A[:,nbase]'*lambda

    #teste
    #println("z = $(dot(c,x))")
    #println("lambda = $lambda")
    #println("sn = $sn")

  end

  #println("Ponto ótimo encontrado!")
  #println("z = $(dot(c,x))")

  return x, base, nbase #retornar ponto otimo
end

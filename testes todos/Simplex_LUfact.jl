function SSolveFact(c,A,b)

  (m,n) = size(A)

  # Fase 1
  ca = [zeros(n);ones(m)]
  B = [A eye(m)]
  xb = b

  for i = 1:m
    if b[i] < 0
      B[i,n+i] = -1
      xb[i] = -b[i]
    end
  end

  #println(xb)
  # resolve o Simplex com variáveis artificiais para achar base inicial
  x, base = Simplex_Rev_Fact(ca,B,b,collect(n+1:n+m),collect(1:n),xb)
  nbase = zeros(Int64,n-m)
  #println(x)
  #println(base)
  #println("nbase = $nbase")

  j = 1
  for i = 1:n
    if all(i.!= base)
      nbase[j] = i
      j = j+1
    end
  end
  #println("nbase = $nbase")
  # resolve o PL original
  x, base, nbase = Simplex_Rev_Fact(c,A,b,base,nbase,x[base])
  return x, base, nbase , dot(c,x)
end


function Simplex_Rev_Fact(c,A,b,base,nbase,xb)

  (m,n) = size(A)
  dimb = length(base)
  dimnb = length(nbase)
  x = zeros(n)
  lambda = zeros(m)

  #println(base)
  #println("nbase = $nbase")

  # encontra primeira fatoração LU com pivotemento
  LU = lufact(sparse(full((A[:,base]))))
  x[base] = xb

  # Resolver o sistema B'v=cb para achar  o vetor v (lambda tamanho m)

  lambda = LU'\(c[base])
  #y = solve_tri_inf(U',c[base])
  #lambda[P] = solve_tri_sup(L', y)

  # Calcular sn = cn-N'*v (tamanho m - n)
  sn = c[nbase] - A[:,nbase]'*lambda

  #teste
  #println(lambda)
  #println(sn)

  #i = 0
  while minimum(sn) < 0.0 && abs(minimum(sn)) > 1e-9#1e-10
    #achar os índices q e p
    q = 0
    for k = 1:dimnb
      if sign(sn[k]) == -1 && abs(sn[k]) > 1e-9#1e-10
  	    q = k
  	    break
      end
    end

    d = LU\(full(A[:,nbase[q]]))
    #y = solve_tri_inf(L,A[P,nbase[q]])
    #d = solve_tri_sup(U, y)

    if  all(d.<= 0)
     error("O problema é ilimitado!")
    end
    alpha = Inf

    p = 0
    for j = 1:dimb
      if d[j] > 1e-12 && x[base[j]] / d[j] < alpha
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

    LU = lufact(sparse(full((A[:,base]))))

    lambda = LU'\c[base]
    #y = solve_tri_inf(U',c[base])
    #lambda[P] = solve_tri_sup(L', y)

    sn = c[nbase] - A[:,nbase]'*lambda
    #println("z = $(dot(c,x))")

  end

  #println("Ponto ótimo encontrado!")
  return x, base, nbase #retornar ponto otimo
end

function exercicio2(A :: SparseMatrixCSC)

  n = size(A, 1)
  U = copy(A)
  L = speye(n)
  P = collect(1:n)
  Q = collect(1:n)

  for k = 1:n-1

    minimo, ki, kj, pivo = Inf, k, k, abs(U[k,k])

    for i = k:n
      for j = k:n
        #(não zeros linha i - 1)*(não zeros coluna j - 1)
        aux = (countnz(U[i,k:n])-1)*(countnz(U[k:n,j])-1)

        if aux < minimo
          if abs(U[i,j]) >= maximum(abs(U[k:n, k:n])) #ou (1/n)*maximum(abs(U[k:n, k:n]))
            # o minimo melhorou e o pivo não piorou
            minimo, ki, kj, pivo = aux, i, j, abs(U[i,j])
          end
        elseif aux == minimo && abs(U[i,j]) > pivo
          # o minimo piorou mas o pivo melhorou
          minimo, ki, kj, pivo = aux, i, j, abs(U[i,j])
        end

      end
    end

    # troca linha ki escolhida com linha k da iteracao
    if ki != k
      I, J = [ki;k], [k;ki]
      P[I] = P[J]
      U[I,k:n] = U[J,k:n]
      L[I,1:k-1] = L[J,1:k-1]
    end

    # troca coluna kj escolhida com coluna k da iteracao
    if kj != k
      I, J = [kj;k], [k;kj]
      Q[I] = Q[J]
      U[:,I] = U[:,J]
    end

    # faz iteracao k do L(PUQ) = U (vetorial)
    #for i = k+1:n
    #  if abs(U[i,k]) > 1e-12
    #    lik = U[i,k]/U[k,k]
    #    L[i,k] = lik
    #    U[i,k] = 0.0
    #    U[i,k+1:n] = U[i,k+1:n] - lik*U[k,k+1:n]
    #  end
    #end
    I = k+1:n
    L[I,k] = U[I,k]/U[k,k]
    U[I,k] = 0
    U[I,I] = U[I,I] - L[I,k] * U[k,I]'
  end

  return L, U, P, Q
end

function exercicio2(A :: AbstractMatrix)

n = size(A, 1)
U = copy(A)
L = eye(n)
P = collect(1:n)
Q = collect(1:n)

for k = 1:n-1

  minimo, ki, kj, pivo = Inf, k, k, abs(U[k,k])

  for i = k:n
    for j = k:n
      #(não zeros linha i - 1)*(não zeros coluna j - 1)
      aux = (countnz(U[i,k:n])-1)*(countnz(U[k:n,j])-1)

      if aux < minimo
        if abs(U[i,j]) >= (1/n)*maximum(abs(U[k:n, j]))
          # o minimo melhorou e o pivo não piorou
          minimo, ki, kj, pivo = aux, i, j, abs(U[i,j])
        end
      elseif aux == minimo && abs(U[i,j]) > pivo
        # o minimo piorou mas o pivo melhorou
        minimo, ki, kj, pivo = aux, i, j, abs(U[i,j])
      end

    end
  end

  # troca linha ki escolhida com linha k da iteracao
  if ki != k
    I, J = [ki;k], [k;ki]
    P[I] = P[J]
    U[I,k:n] = U[J,k:n]
    L[I,1:k-1] = L[J,1:k-1]
  end

  # troca coluna kj escolhida com coluna k da iteracao
  if kj != k
    I, J = [kj;k], [k;kj]
    Q[I] = Q[J]
    U[:,I] = U[:,J]
  end

  # faz iteracao k do L(PUQ) = U (vetorial)
  I = k+1:n
  L[I,k] = U[I,k]/U[k,k]
  U[I,k] = 0
  U[I,I] = U[I,I] - L[I,k] * U[k,I]'
end

return L, U, P, Q
end

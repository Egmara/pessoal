# LU update primeira tentativa

function LUupdate(U,r,p,m,w)

  R = copy(U)

  # vetor que representa matriz P transposta da teoria (U*P')
  P = [collect(1:p-1);collect(p+1:r);p;collect(r+1:m)]
  # vetor que guarda a linha densa de Ltiu inversa
  l = zeros(m)
  l[r] = 1

  # spike
  R[:,p] = w
  # permuta colunas 
  R = R[:,P]

  # zera elementos na linha p da coluna p até r-1
  for j = p:r-1
    # zera se for necessário
    if abs(R[p,j]) > 1e-10 #0.0
      escalar = R[p,j]/R[j+1,j]
      # guarda a inversa
      l[j] = -escalar
      # atualiza linha p
      R[p,j] = 0.0
      R[p,j+1:m] = R[p,j+1:m] - escalar*R[j+1,j+1:m]
    end
  end

  # permuta linhas
  R = R[P,:]

  return R,l,P

end

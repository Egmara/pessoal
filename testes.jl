using Clp, JuMP, MathProgBase

mod=Model(solver=ClpSolver(PreCrush=1, Cuts=0, Presolve=0, Heuristics=0.0, DisplayInterval=1)) #tipo do modelo
m_internal = MathProgBase.LinearQuadraticModel(ClpSolver()) #tipo do solver
#include("testmps.jl")
elementos=["aa01.mps";"aa03.mps";"aa3.mps";"aa4.mps";"aa5.mps";"aa6.mps";"air02.mps";"air03.mps";"air04.mps";"air05.mps";"air06.mps";"aircraft.mps";"bas1lp.mps";"baxter.mps";"car4.mps";"cari.mps";"ch.mps";"co5.mps";"co9.mps";"complex.mps";"cq5.mps";"cq9.mps";"cr42.mps";"crew1.mps";"dano3mip.mps";"dbic1.mps";"dbir1.mps";"dbir2.mps";"delf000.mps";"delf001.mps";"delf002.mps";"delf003.mps";"delf004.mps";"delf005.mps";"delf006.mps";"delf007.mps";"delf008.mps";"delf009.mps";"delf010.mps";"delf011.mps";"delf012.mps";"delf013.mps";"delf014.mps";"delf015.mps";"delf017.mps";"delf018.mps";"delf019.mps";"delf020.mps";"delf021.mps";"delf022.mps";"delf023.mps";"delf024.mps";"delf025.mps";"delf026.mps";"delf027.mps";"delf028.mps";"delf029.mps";"delf030.mps";"delf031.mps";"delf032.mps";"delf033.mps";"delf034.mps";"delf035.mps";"delf036.mps";"df2177.mps";"disp3.mps";"dsbmip.mps";"e18.mps";"ex3sta1.mps";"farm.mps";"gams10a.mps";"gams30a.mps";"ge.mps";"iiasa.mps";"jendrec1.mps";"kent.mps";"kl02.mps";"kleemin3.mps";"kleemin4.mps";"kleemin5.mps";"kleemin6.mps";"kleemin7.mps";"kleemin8.mps";"l9.mps";"large000.mps";"large001.mps";"large002.mps";"large003.mps";"large004.mps";"large005.mps";"large006.mps";"large007.mps";"large008.mps";"large009.mps";"large010.mps";"large011.mps";"large012.mps";"large013.mps";"large014.mps";"large015.mps";"large016.mps";"large017.mps";"large018.mps";"large019.mps";"large020.mps";"large021.mps";"large022.mps";"large023.mps";"large024.mps";"large025.mps";"large026.mps";"large027.mps";"large028.mps";"large029.mps";"large030.mps";"large031.mps";"large032.mps";"large033.mps";"large034.mps";"large035.mps";"large036.mps";"lp22.mps";"lpl1.mps";"lpl2.mps";"lpl3.mps";"mod2.mps";"model1.mps";"model2.mps";"model3.mps";"model4.mps";"model5.mps";"model6.mps";"model7.mps";"model8.mps";"model9.mps";"model10.mps";"model11.mps";"multi.mps";"nemsafm.mps";"nemscem.mps";"nemsemm1.mps";"nemsemm2.mps";"nemspmm1.mps";"nemspmm2.mps";"nemswrld.mps";"nl.mps";"nsct1.mps";"nsct2.mps";"nsic1.mps";"nsic2.mps";"nsir1.mps";"nsir2.mps";"nug05.mps";"nug06.mps";"nug07.mps";"nug08.mps";"nug12.mps";"nug15.mps";"nw14.mps";"orna1.mps";"orna2.mps";"orna3.mps";"orna4.mps";"orna7.mps";"orswq2.mps";"p0033.mps";"p0040.mps";"p010.mps";"p0201.mps";"p0282.mps";"p0291.mps";"p05.mps";"p0548.mps";"p19.mps";"p2756.mps";"p6000.mps";"pcb1000.mps";"pcb3000.mps";"pf2177.mps";"pldd000b.mps";"pldd001b.mps";"pldd002b.mps";"pldd003b.mps";"pldd004b.mps";"pldd005b.mps";"pldd006b.mps";"pldd007b.mps";"pldd008b.mps";"pldd009b.mps";"pldd010b.mps";"pldd011b.mps";"pldd012b.mps";"primagaz.mps";"problem.mps";"progas.mps";"qiulp.mps";"r05.mps";"rat1.mps";"rat5.mps";"rat7a.mps";"refine.mps";"rlfddd.mps";"rlfdual.mps";"rlfprim.mps";"rosen1.mps";"rosen2.mps";"rosen7.mps";"rosen8.mps";"rosen10.mps";"route.mps";"seymourl.mps";"slptsk.mps";"small000.mps";"small001.mps";"small002.mps";"small003.mps";"small004.mps";"small005.mps";"small006.mps";"small007.mps";"small008.mps";"small009.mps";"small010.mps";"small011.mps";"small012.mps";"small013.mps";"small014.mps";"small015.mps";"small016.mps";"south31.mps";"stat96v1.mps";"stat96v2.mps";"stat96v3.mps";"stat96v4.mps";"stat96v5.mps";"sws.mps";"t0331-4l.mps";"testbig.mps";"ulevimin.mps";"us04.mps";"world.mps";"zed.mps"]
for fun=1:length(elementos) #para cada funçao
    MathProgBase.loadproblem!(m_internal, elementos[fun]) #carrega os dados da função

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
    println(elementos[fun])
  end

end

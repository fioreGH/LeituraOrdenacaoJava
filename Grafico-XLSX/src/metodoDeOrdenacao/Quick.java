package metodoDeOrdenacao;

public class Quick {
	
	public int[] Quick(int[] v, int inicio, int fim) {
		if (inicio < fim) {
			int posicaoPivo = separar(v, inicio, fim);
			Quick(v, inicio, posicaoPivo - 1);
			Quick(v, posicaoPivo + 1, fim);
		}
		return v;
	}
	

	public int separar(int[] vetor, int inicio, int fim) {
	    int pivo = vetor[inicio];
	    int i = inicio + 1, f = fim;
	    while (i <= f) {
	           if (vetor[i] <= pivo)
	                  i++;
	           else if (pivo < vetor[f])
	                  f--;
	           else {
	                  int troca = vetor[i];
	                  vetor[i] = vetor[f];
	                  vetor[f] = troca;
	                  i++;
	                  f--;
	           }
	    }
	    vetor[inicio] = vetor[f];
	    vetor[f] = pivo;
	    return f;
	}
}

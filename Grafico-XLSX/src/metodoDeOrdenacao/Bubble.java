package metodoDeOrdenacao;

public class Bubble {
	
	public int[] Bubble(int[] v) {
		
		int i = 0;
		int j = 0;
		int k = 0;
		int auxiliar = 0;
		int tamanho = v.length;
		
		for (i = 0; i < tamanho - 1; i++) {
	        for (j = 0; j < tamanho - 1 - i; j++) {
	            if (v[j] > v[j + 1]) {
	                auxiliar = v[j];
	                v[j] = v[j + 1];
	                v[j + 1] = auxiliar;
	            }
	        }
	    }
		
	    return v;
		
	}

}

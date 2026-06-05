#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <time.h>

typedef struct {
    int awb;
    char produto[50];
    float valor;
    char status[20];
} Pacote;

typedef struct {
    int id_motorista;
    char nome[50];
    char telefone[15];
    char base[50];
} Motorista;

typedef struct {
    int id_cliente;
    char nome[50];
    char telefone[15];
    char endereco[100];
} Cliente;

typedef struct {
    int id_ticket;
    int awb_pacote;
    int id_cliente;
    int id_motorista;
    char prazo[15];
    char status[20];
} Ticket;

void gerar_base_pacotes(int tamanho) {
    FILE *arq = fopen("pacotes.dat", "wb");
    if (arq == NULL) {
        printf("Erro ao criar arquivo de pacotes.\n");
        return;
    }
    
    for (int i = 1; i <= tamanho; i++) {
        Pacote p;
        p.awb = i;
        sprintf(p.produto, "Produto_Generico_%d", i);
        p.valor = 10.0 * i;
        strcpy(p.status, "Em Rota");
        fwrite(&p, sizeof(Pacote), 1, arq);
    }
    fclose(arq);
    printf("Base com %d pacotes gerada com sucesso!\n", tamanho);
}

void registrar_log(const char *metodo, int tamanho, int comparacoes, double tempo) {
    FILE *log = fopen("log.txt", "a");
    if (log != NULL) {
        fprintf(log, "Metodo: %-15s | Tamanho Base: %-7d | Comparacoes: %-6d | Tempo: %f seg\n", 
                metodo, tamanho, comparacoes, tempo);
        fclose(log);
    }
}

Pacote* busca_sequencial_pacote(int awb_buscado, int tamanho_base) {
    FILE *arq = fopen("pacotes.dat", "rb");
    if (arq == NULL) return NULL;

    Pacote *p = malloc(sizeof(Pacote));
    int comparacoes = 0;
    
    clock_t inicio_t = clock();

    while (fread(p, sizeof(Pacote), 1, arq) == 1) {
        comparacoes++;
        if (p->awb == awb_buscado) {
            break;
        }
    }

    clock_t fim_t = clock();
    double tempo_gasto = (double)(fim_t - inicio_t) / CLOCKS_PER_SEC;

    registrar_log("Sequencial", tamanho_base, comparacoes, tempo_gasto);
    fclose(arq);
    
    if (p->awb == awb_buscado) return p;
    free(p);
    return NULL;
}

Pacote* busca_binaria_pacote(int awb_buscado, int tamanho_base) {
    FILE *arq = fopen("pacotes.dat", "rb");
    if (arq == NULL) return NULL;

    Pacote *p = malloc(sizeof(Pacote));
    int inicio = 0;
    int fim = tamanho_base - 1;
    int comparacoes = 0;
    int encontrou = 0;

    clock_t inicio_t = clock();

    while (inicio <= fim) {
        int meio = (inicio + fim) / 2;
        
        fseek(arq, meio * sizeof(Pacote), SEEK_SET); 
        fread(p, sizeof(Pacote), 1, arq);
        
        comparacoes++;
        
        if (p->awb == awb_buscado) {
            encontrou = 1;
            break;
        } else if (p->awb < awb_buscado) {
            inicio = meio + 1;
        } else {
            fim = meio - 1;
        }
    }

    clock_t fim_t = clock();
    double tempo_gasto = (double)(fim_t - inicio_t) / CLOCKS_PER_SEC;

    registrar_log("Binaria", tamanho_base, comparacoes, tempo_gasto);
    fclose(arq);

    if (encontrou) return p;
    free(p);
    return NULL;
}

void mostrar_base_pacotes() {
    FILE *arq = fopen("pacotes.dat", "rb");
    if (arq == NULL) {
        printf("Nenhum pacote registrado ou erro ao abrir o arquivo.\n");
        return;
    }
    
    Pacote p;
    int contador = 0;
    int limite = 1000000;

    printf("\n--- BASE DE DADOS DE PACOTES (Mostrando primeiros %d) ---\n", limite);
    printf("%-10s | %-25s | %-10s | %-15s\n", "AWB", "PRODUTO", "VALOR (R$)", "STATUS");
    printf("-----------------------------------------------------------------------\n");

    while (fread(&p, sizeof(Pacote), 1, arq) == 1) {
        printf("%-10d | %-25s | %-10.2f | %-15s\n", 
               p.awb, p.produto, p.valor, p.status);
        
        contador++;
        if (contador >= limite) {
            printf("... (Exibição limitada aos primeiros %d pacotes) ...\n", limite);
            break;
        }
    }
    
    fclose(arq);
    printf("-----------------------------------------------------------------------\n");
}

void abrir_ticket_acareacao(int awb, int id_cli, int id_mot, char *prazo, int tamanho_base) {
    printf("\n--- ABRINDO TICKET DE ACAREACAO ---\n");
    
    Pacote *p = busca_binaria_pacote(awb, tamanho_base);
    
    if (p != NULL) {
        Ticket t;
        t.id_ticket = rand() % 1000;
        t.awb_pacote = p->awb;
        t.id_cliente = id_cli;
        t.id_motorista = id_mot;
        strcpy(t.prazo, prazo);
        strcpy(t.status, "Pendente");

        FILE *arq = fopen("tickets.dat", "ab");
        fwrite(&t, sizeof(Ticket), 1, arq);
        fclose(arq);

        printf("Ticket ID %d aberto com sucesso para o AWB %d (%s)!\n", t.id_ticket, p->awb, p->produto);
        free(p);
    } else {
        printf("Falha: Pacote AWB %d nao encontrado na base!\n", awb);
    }
}

void listar_tickets_pendentes() {
    FILE *arq = fopen("tickets.dat", "rb");
    if (arq == NULL) {
        printf("Nenhum ticket registrado.\n");
        return;
    }
    
    Ticket t;
    printf("\n--- TICKETS PENDENTES ---\n");
    while (fread(&t, sizeof(Ticket), 1, arq) == 1) {
        if (strcmp(t.status, "Pendente") == 0) {
            printf("ID Ticket: %d | AWB: %d | Motorista: %d | Prazo: %s\n", 
                   t.id_ticket, t.awb_pacote, t.id_motorista, t.prazo);
        }
    }
    fclose(arq);
}

int main() {
    int opcao;
    int tamanho_atual = 10000;
    
    remove("log.txt");

    do {
        printf("\n==========================================\n");
        printf(" SISTEMA DE ACAREACOES LOGISTICAS\n");
        printf("==========================================\n");
        printf("1. Gerar Base de Pacotes (Tamanho: %d)\n", tamanho_atual);
        printf("2. Mudar Tamanho da Base\n");
        printf("3. Testar Busca Sequencial\n");
        printf("4. Testar Busca Binaria\n");
        printf("5. Operacao: Abrir Novo Ticket\n");
        printf("6. Operacao: Listar Tickets Pendentes\n");
        printf("7. Mostrar Base de Pacotes\n");
        printf("0. Sair\n");
        printf("Escolha: ");
        scanf("%d", &opcao);

        if (opcao == 1) {
            gerar_base_pacotes(tamanho_atual);
        } 
        else if (opcao == 2) {
            printf("Digite o novo tamanho (ex: 1000, 50000): ");
            scanf("%d", &tamanho_atual);
            gerar_base_pacotes(tamanho_atual);
        }
        else if (opcao == 3 || opcao == 4) {
            int awb_teste;
            printf("Digite o AWB a buscar (1 a %d): ", tamanho_atual);
            scanf("%d", &awb_teste);
            
            Pacote *resultado = NULL;
            if (opcao == 3) {
                printf("\nExecutando Busca Sequencial...\n");
                resultado = busca_sequencial_pacote(awb_teste, tamanho_atual);
            } else {
                printf("\nExecutando Busca Binaria...\n");
                resultado = busca_binaria_pacote(awb_teste, tamanho_atual);
            }
            
            if (resultado != NULL) {
                printf("Pacote Encontrado! Produto: %s | Valor: %.2f | Status: %s\n", 
                       resultado->produto, resultado->valor, resultado->status);
                free(resultado);
            } else {
                printf("Pacote nao encontrado.\n");
            }
            printf("Verifique o arquivo 'log.txt' para os dados de performance.\n");
        }
        else if (opcao == 5) {
            int awb, cli, mot;
            char prazo[15];
            printf("AWB do Pacote reclamado: "); scanf("%d", &awb);
            printf("ID do Cliente: "); scanf("%d", &cli);
            printf("ID do Motorista: "); scanf("%d", &mot);
            printf("Prazo Limite (DD/MM/AAAA): "); scanf("%s", prazo);
            abrir_ticket_acareacao(awb, cli, mot, prazo, tamanho_atual);
        }
        else if (opcao == 6) {
            listar_tickets_pendentes();
        }
        else if (opcao == 7) {
            mostrar_base_pacotes();
        }

    } while (opcao != 0);

    return 0;
}
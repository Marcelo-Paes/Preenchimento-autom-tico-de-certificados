# Preenchimento-automático-de-certificados
Este aplicativo automatiza o processo de preenchimento de informações em certificados utilizando dados de uma planilha. Aqui está como ele funciona:

1. **Entrada de Dados**: O usuário fornece uma planilha (no formato xlsx) contendo informações como o nome do curso, nome do participante, tipo de participação, datas de início e término, carga horária, data de emissão do certificado, e possivelmente assinaturas.

2. **Extração de Dados**: O programa lê os dados da planilha, linha por linha, e extrai as informações necessárias para preencher o certificado.

3. **Transferência para a Imagem do Certificado**: Utilizando a biblioteca PIL (Python Imaging Library), o aplicativo abre um modelo de certificado (provavelmente em formato de imagem), e desenha as informações extraídas nos locais apropriados do certificado.

4. **Personalização e Salvar**: Depois de inserir todas as informações no certificado, ele é salvo em um arquivo de imagem (como PNG) com um nome específico, possivelmente contendo o nome do participante para identificação.

5. **Repetição para Cada Participante**: O processo é repetido para cada linha na planilha, criando um certificado personalizado para cada participante listado.

Este aplicativo economiza tempo e esforço, eliminando a necessidade de preencher manualmente cada certificado com as informações dos participantes. Ele é útil em situações onde há um grande número de certificados para emitir, como em eventos, cursos ou workshops.

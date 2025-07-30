{
 "cells": [
  {
   "cell_type": "markdown",
   "id": "ef349ab6-aad3-4625-b773-143922a44106",
   "metadata": {},
   "source": [
    "# PD EM MASSA"
   ]
  },
  {
   "cell_type": "markdown",
   "id": "4b27c35d-db4a-426f-8c8b-22afd9e80ad8",
   "metadata": {},
   "source": [
    "## Imports"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 2,
   "id": "dc5efd47-eeef-48ef-b063-0d76ee938a07",
   "metadata": {},
   "outputs": [],
   "source": [
    "import requests\n",
    "import pandas as pd\n",
    "from pathlib import Path\n",
    "from datetime import datetime\n",
    "from io import StringIO\n",
    "from openpyxl import load_workbook\n",
    "import os"
   ]
  },
  {
   "cell_type": "markdown",
   "id": "917ff7f5-de9f-4a1f-9022-ca07b0083ce0",
   "metadata": {},
   "source": [
    "## Components"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 5,
   "id": "7ffe7f41-3942-4e07-b265-a00fe1c0f103",
   "metadata": {},
   "outputs": [],
   "source": [
    "#base_dir = Path(r\"C:\\Users\\Mateus.Martins\\supremocimento.com.br\\[CORP] Inteligencia de Mercado - Documentos\\General\\13. Projetos\\9. Pricing\\General\\13. Projetos\\9. Pricing\\1.PD_em_massa\\4. PD_em_massa_auto\")\n",
    "#export_dir = base_dir / \"Save\"\n",
    "# export_dir.mkdir(parents=True, exist_ok=True)\n",
    "# aguardando_analise_save = base_dir / \"2.aguardando_analise\" \n",
    "# aguardando_analise_save.mkdir(parents=True, exist_ok=True)"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 9,
   "id": "8ed8a174-a965-463c-8932-e687d9b8ee0b",
   "metadata": {},
   "outputs": [],
   "source": [
    "base_dir = Path.cwd()\n",
    "export_dir = base_dir / \"Save\""
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 30,
   "id": "e9420d71-6900-417b-a38a-7b80a7aef37e",
   "metadata": {},
   "outputs": [],
   "source": [
    "# Cookie SID (copiado da sessão ativa do navegador)\n",
    "\n",
    "#edge://settings/cookies/detail?site=secil.my.salesforce.com\n",
    "\n",
    "sid = \"00DD0000000maML!AQEAQFZpTwW9nreIIVwd2xDeqYU.rkqchKnJBeoZebu2E91N17s3ZxecVY8MlyS9HvKWgSs5ivPqD_7BZMfkTTOefthSfLLJ\""
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 32,
   "id": "89b52273-2861-4141-be77-6a42bad89e17",
   "metadata": {},
   "outputs": [],
   "source": [
    "# Cabeçalhos HTTP com o cookie\n",
    "headers = {\n",
    "    \"Cookie\": f\"sid={sid}\",\n",
    "    \"User-Agent\": \"Mozilla/5.0\"\n",
    "}"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 34,
   "id": "11415753-0abe-4d01-b461-5632ae79583b",
   "metadata": {},
   "outputs": [],
   "source": [
    "# Data de hoje formatada\n",
    "data_hoje = datetime.now().strftime(\"%d-%m-%Y\")"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 36,
   "id": "31e2d88e-5370-4324-81aa-97e2b22ded26",
   "metadata": {},
   "outputs": [],
   "source": [
    "relatorios = {\n",
    "    \"pd_massa\": \"https://secil.my.salesforce.com/00O7S000001kByi?export=1&enc=UTF-8&xf=csv\",\n",
    "}"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 38,
   "id": "80f6bd55-90aa-4a3c-8938-986bff24fdbe",
   "metadata": {},
   "outputs": [],
   "source": [
    "# Lista de colunas que devem ser convertidas para números (float)\n",
    "colunas_numericas = [\n",
    "    'Preço Tabela / Saco',\n",
    "    'Preço Proposto FOB /Ton',\n",
    "    'Preço Proposto Frete /Ton',\n",
    "    'Preço Proposto Final /Ton',\n",
    "    'Valor do frete agenciado',\n",
    "    'Preço Proposto FOB /Saco',\n",
    "    'Preço Proposto Frete /Saco',\n",
    "    'Preço Proposto Final /Saco'\n",
    "]"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 40,
   "id": "7e6fee02-6a74-4c2f-9836-bca16cb22b60",
   "metadata": {},
   "outputs": [],
   "source": [
    "# Lista de nomes que você deseja extrair\n",
    "responsaveis = {\n",
    "    'Alessandro Araujo': 'Mateus',\n",
    "    'Allan Andrade': 'Mateus',\n",
    "    'Andrew Bairros': 'Andrew',\n",
    "    'Andrey Monteiro': 'Andrey',\n",
    "    'Asaph Nascimento': 'Mateus',\n",
    "    'Claudemir Muller': 'Andrew',\n",
    "    'Edilson Boron': 'Mateus',\n",
    "    'Eduardo Bryk': 'Mateus',\n",
    "    'Fabio Pedrini': 'Thiago',\n",
    "    'Gilmar Jesus': 'Andrey',\n",
    "    'Giovani Nogueira': 'Thiago',\n",
    "    'Giovany Borsoi': 'Thiago',\n",
    "    'Giulliano Oliveira': 'Andrew',\n",
    "    'Jordana Porto': 'Andrew',\n",
    "    'José Ferreira': 'Andrey',\n",
    "    'Juliano Rezende': 'Mateus',\n",
    "    'Juliano Scherer': 'Andrew',\n",
    "    'Leandro Ceron': 'Thiago',\n",
    "    'Lucas Alves': 'Thiago',\n",
    "    'Mateus Antonangelo': 'Mateus',\n",
    "    'Thiago Bergara': 'Andrey',\n",
    "    'Thiago Senise': 'Thiago',\n",
    "    'Vagner Lopes': 'Mateus',\n",
    "    'Valdete Pilar': 'Thiago',\n",
    "    'Weverson Raimundo': 'Mateus'\n",
    "}"
   ]
  },
  {
   "cell_type": "markdown",
   "id": "f4aead1b-6b4d-4624-89b0-576038beb375",
   "metadata": {},
   "source": [
    "## Code"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 43,
   "id": "9d8571bf-0a1b-4f5d-a2f4-f9d24536785b",
   "metadata": {},
   "outputs": [],
   "source": [
    "# Função de limpeza numérica\n",
    "def para_float(coluna):\n",
    "    return (\n",
    "        coluna.astype(str)\n",
    "        .str.replace(\",\", \".\", regex=False)\n",
    "        .str.replace(\"%\", \"\", regex=False)\n",
    "        .str.strip()\n",
    "    )"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 45,
   "id": "b5faaa2d-f5e6-4fde-91ed-1c78856829cf",
   "metadata": {
    "scrolled": true
   },
   "outputs": [
    {
     "name": "stdout",
     "output_type": "stream",
     "text": [
      "🔄 Baixando pd_massa...\n",
      "🗑️ Apagado: PDsAlessandro Araujo(19-05-2025).xlsx\n",
      "✅ Salvo: C:\\Users\\Mateus.Martins\\supremocimento.com.br\\[POM] Comercial - Documentos\\General\\IM\\PD em massa\\Mateus\\Alessandro Araujo\\PDsAlessandro Araujo(30-05-2025).xlsx\n",
      "🗑️ Apagado: PDsAllan Andrade(19-05-2025).xlsx\n",
      "✅ Salvo: C:\\Users\\Mateus.Martins\\supremocimento.com.br\\[POM] Comercial - Documentos\\General\\IM\\PD em massa\\Mateus\\Allan Andrade\\PDsAllan Andrade(30-05-2025).xlsx\n",
      "🗑️ Apagado: PDsAndrew Bairros(19-05-2025).xlsx\n",
      "✅ Salvo: C:\\Users\\Mateus.Martins\\supremocimento.com.br\\[POM] Comercial - Documentos\\General\\IM\\PD em massa\\Andrew\\Andrew Bairros\\PDsAndrew Bairros(30-05-2025).xlsx\n",
      "🗑️ Apagado: PDsAndrey Monteiro(19-05-2025).xlsx\n",
      "✅ Salvo: C:\\Users\\Mateus.Martins\\supremocimento.com.br\\[POM] Comercial - Documentos\\General\\IM\\PD em massa\\Andrey\\Andrey Monteiro\\PDsAndrey Monteiro(30-05-2025).xlsx\n",
      "🗑️ Apagado: PDsAsaph Nascimento(19-05-2025).xlsx\n",
      "✅ Salvo: C:\\Users\\Mateus.Martins\\supremocimento.com.br\\[POM] Comercial - Documentos\\General\\IM\\PD em massa\\Mateus\\Asaph Nascimento\\PDsAsaph Nascimento(30-05-2025).xlsx\n",
      "🗑️ Apagado: PDsClaudemir Muller(19-05-2025).xlsx\n",
      "✅ Salvo: C:\\Users\\Mateus.Martins\\supremocimento.com.br\\[POM] Comercial - Documentos\\General\\IM\\PD em massa\\Andrew\\Claudemir Muller\\PDsClaudemir Muller(30-05-2025).xlsx\n",
      "🗑️ Apagado: PDsEdilson Boron(19-05-2025).xlsx\n",
      "✅ Salvo: C:\\Users\\Mateus.Martins\\supremocimento.com.br\\[POM] Comercial - Documentos\\General\\IM\\PD em massa\\Mateus\\Edilson Boron\\PDsEdilson Boron(30-05-2025).xlsx\n",
      "🗑️ Apagado: PDsEduardo Bryk(19-05-2025).xlsx\n",
      "✅ Salvo: C:\\Users\\Mateus.Martins\\supremocimento.com.br\\[POM] Comercial - Documentos\\General\\IM\\PD em massa\\Mateus\\Eduardo Bryk\\PDsEduardo Bryk(30-05-2025).xlsx\n",
      "🗑️ Apagado: PDsFabio Pedrini(19-05-2025).xlsx\n",
      "✅ Salvo: C:\\Users\\Mateus.Martins\\supremocimento.com.br\\[POM] Comercial - Documentos\\General\\IM\\PD em massa\\Thiago\\Fabio Pedrini\\PDsFabio Pedrini(30-05-2025).xlsx\n",
      "🗑️ Apagado: PDsGilmar Jesus(19-05-2025).xlsx\n",
      "✅ Salvo: C:\\Users\\Mateus.Martins\\supremocimento.com.br\\[POM] Comercial - Documentos\\General\\IM\\PD em massa\\Andrey\\Gilmar Jesus\\PDsGilmar Jesus(30-05-2025).xlsx\n",
      "🗑️ Apagado: PDsGiovani Nogueira(19-05-2025).xlsx\n",
      "✅ Salvo: C:\\Users\\Mateus.Martins\\supremocimento.com.br\\[POM] Comercial - Documentos\\General\\IM\\PD em massa\\Thiago\\Giovani Nogueira\\PDsGiovani Nogueira(30-05-2025).xlsx\n",
      "🗑️ Apagado: PDsGiovany Borsoi(19-05-2025).xlsx\n",
      "✅ Salvo: C:\\Users\\Mateus.Martins\\supremocimento.com.br\\[POM] Comercial - Documentos\\General\\IM\\PD em massa\\Thiago\\Giovany Borsoi\\PDsGiovany Borsoi(30-05-2025).xlsx\n",
      "🗑️ Apagado: PDsGiulliano Oliveira(19-05-2025).xlsx\n",
      "✅ Salvo: C:\\Users\\Mateus.Martins\\supremocimento.com.br\\[POM] Comercial - Documentos\\General\\IM\\PD em massa\\Andrew\\Giulliano Oliveira\\PDsGiulliano Oliveira(30-05-2025).xlsx\n",
      "🗑️ Apagado: PDsJordana Porto(19-05-2025).xlsx\n",
      "✅ Salvo: C:\\Users\\Mateus.Martins\\supremocimento.com.br\\[POM] Comercial - Documentos\\General\\IM\\PD em massa\\Andrew\\Jordana Porto\\PDsJordana Porto(30-05-2025).xlsx\n",
      "🗑️ Apagado: PDsJosé Ferreira(19-05-2025).xlsx\n",
      "✅ Salvo: C:\\Users\\Mateus.Martins\\supremocimento.com.br\\[POM] Comercial - Documentos\\General\\IM\\PD em massa\\Andrey\\José Ferreira\\PDsJosé Ferreira(30-05-2025).xlsx\n",
      "🗑️ Apagado: PDsJuliano Rezende(19-05-2025).xlsx\n",
      "✅ Salvo: C:\\Users\\Mateus.Martins\\supremocimento.com.br\\[POM] Comercial - Documentos\\General\\IM\\PD em massa\\Mateus\\Juliano Rezende\\PDsJuliano Rezende(30-05-2025).xlsx\n",
      "🗑️ Apagado: PDsJuliano Scherer(19-05-2025).xlsx\n",
      "✅ Salvo: C:\\Users\\Mateus.Martins\\supremocimento.com.br\\[POM] Comercial - Documentos\\General\\IM\\PD em massa\\Andrew\\Juliano Scherer\\PDsJuliano Scherer(30-05-2025).xlsx\n",
      "🗑️ Apagado: PDsLeandro Ceron(19-05-2025).xlsx\n",
      "✅ Salvo: C:\\Users\\Mateus.Martins\\supremocimento.com.br\\[POM] Comercial - Documentos\\General\\IM\\PD em massa\\Thiago\\Leandro Ceron\\PDsLeandro Ceron(30-05-2025).xlsx\n",
      "🗑️ Apagado: PDsLucas Alves(19-05-2025).xlsx\n",
      "✅ Salvo: C:\\Users\\Mateus.Martins\\supremocimento.com.br\\[POM] Comercial - Documentos\\General\\IM\\PD em massa\\Thiago\\Lucas Alves\\PDsLucas Alves(30-05-2025).xlsx\n",
      "🗑️ Apagado: PDsMateus Antonangelo(19-05-2025).xlsx\n",
      "✅ Salvo: C:\\Users\\Mateus.Martins\\supremocimento.com.br\\[POM] Comercial - Documentos\\General\\IM\\PD em massa\\Mateus\\Mateus Antonangelo\\PDsMateus Antonangelo(30-05-2025).xlsx\n",
      "🗑️ Apagado: PDsThiago Bergara(19-05-2025).xlsx\n",
      "✅ Salvo: C:\\Users\\Mateus.Martins\\supremocimento.com.br\\[POM] Comercial - Documentos\\General\\IM\\PD em massa\\Andrey\\Thiago Bergara\\PDsThiago Bergara(30-05-2025).xlsx\n",
      "🗑️ Apagado: PDsThiago Senise(19-05-2025).xlsx\n",
      "✅ Salvo: C:\\Users\\Mateus.Martins\\supremocimento.com.br\\[POM] Comercial - Documentos\\General\\IM\\PD em massa\\Thiago\\Thiago Senise\\PDsThiago Senise(30-05-2025).xlsx\n",
      "🗑️ Apagado: PDsVagner Lopes(19-05-2025).xlsx\n",
      "✅ Salvo: C:\\Users\\Mateus.Martins\\supremocimento.com.br\\[POM] Comercial - Documentos\\General\\IM\\PD em massa\\Mateus\\Vagner Lopes\\PDsVagner Lopes(30-05-2025).xlsx\n",
      "🗑️ Apagado: PDsValdete Pilar(19-05-2025).xlsx\n",
      "✅ Salvo: C:\\Users\\Mateus.Martins\\supremocimento.com.br\\[POM] Comercial - Documentos\\General\\IM\\PD em massa\\Thiago\\Valdete Pilar\\PDsValdete Pilar(30-05-2025).xlsx\n",
      "🗑️ Apagado: PDsWeverson Raimundo(19-05-2025).xlsx\n",
      "✅ Salvo: C:\\Users\\Mateus.Martins\\supremocimento.com.br\\[POM] Comercial - Documentos\\General\\IM\\PD em massa\\Mateus\\Weverson Raimundo\\PDsWeverson Raimundo(30-05-2025).xlsx\n",
      " ✅✅✅Processamento concluído às 10:58:04!\n"
     ]
    }
   ],
   "source": [
    "# Download e processamento\n",
    "for nome, url in relatorios.items():\n",
    "    print(f\"🔄 Baixando {nome}...\")\n",
    "    response = requests.get(url, headers=headers)\n",
    "    response.raise_for_status()\n",
    "\n",
    "    if \"html\" in response.text[:100].lower():\n",
    "        raise ValueError(\"⚠️ Resposta parece ser HTML (erro no SID ou link).\")\n",
    "\n",
    "    df = pd.read_csv(StringIO(response.text), on_bad_lines='skip')\n",
    "\n",
    "    for col in colunas_numericas:\n",
    "        if col in df.columns:\n",
    "            df[col] = para_float(df[col])\n",
    "            df[col] = pd.to_numeric(df[col], errors=\"coerce\")\n",
    "\n",
    "    # Salva planilha geral\n",
    "    df.to_excel(export_dir / f\"{nome}({data_hoje}).xlsx\", index=False)\n",
    "\n",
    "    # Geração de arquivos individuais\n",
    "    col_prop = \"Proprietário da conta\"\n",
    "    for prop, pasta in responsaveis.items():\n",
    "        if col_prop not in df.columns:\n",
    "            print(f\"⚠️ Coluna '{col_prop}' não encontrada no relatório.\")\n",
    "            break\n",
    "\n",
    "        if prop not in df[col_prop].values:\n",
    "            print(f\"❌ Nome não encontrado na base: {prop}\")\n",
    "            continue\n",
    "\n",
    "        df_filtrado = df[df[col_prop] == prop]\n",
    "        destino = base_dir / pasta / prop\n",
    "\n",
    "        if not destino.exists():\n",
    "            print(f\"📂 Pasta não encontrada: {destino}\")\n",
    "            continue\n",
    "\n",
    "         #Remove arquivos existentes na pasta do colaborador\n",
    "        for file in destino.glob(\"*\"):\n",
    "            if file.is_file():\n",
    "                try:                     \n",
    "                    file.unlink()\n",
    "                    print(f\"🗑️ Apagado: {file.name}\")\n",
    "                except Exception as e:\n",
    "                    print(f\"⚠️ Erro ao apagar {file.name}: {e}\")\n",
    "\n",
    "        caminho_arquivo = destino / f\"PDs{prop}({data_hoje}).xlsx\"\n",
    "        df_filtrado.to_excel(caminho_arquivo, index=False)\n",
    "        print(f\"✅ Salvo: {caminho_arquivo}\")\n",
    "\n",
    "print(f\" ✅✅✅Processamento concluído às {datetime.now().strftime('%H:%M:%S')}!\")"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": null,
   "id": "8ea71e58-a16e-4ec2-9e18-2212ceb75421",
   "metadata": {},
   "outputs": [],
   "source": []
  }
 ],
 "metadata": {
  "kernelspec": {
   "display_name": "Python 3 (ipykernel)",
   "language": "python",
   "name": "python3"
  },
  "language_info": {
   "codemirror_mode": {
    "name": "ipython",
    "version": 3
   },
   "file_extension": ".py",
   "mimetype": "text/x-python",
   "name": "python",
   "nbconvert_exporter": "python",
   "pygments_lexer": "ipython3",
   "version": "3.12.4"
  }
 },
 "nbformat": 4,
 "nbformat_minor": 5
}

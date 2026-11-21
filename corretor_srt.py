import importlib
import json
import re
import sqlite3
import threading
import time
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Protocol
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import urlopen


MARKUP_TAG_PATTERN = re.compile(
    r"</?[A-Za-z][A-Za-z0-9:-]*(?:\s+[^<>]*)?>", re.IGNORECASE
)
BRACKETED_DESCRIPTION_PATTERN = re.compile(r"\[[^\]]*\]")
PROMOTIONAL_TEXT_PATTERN = re.compile(r"SimpleFileTranslator", re.IGNORECASE)


class TranslationError(RuntimeError):
    pass


class ArgosTranslation(Protocol):
    def translate(self, text: str) -> str: ...


TRANSLATION_CACHE = Path.home() / ".legendas_corretor" / "translation_cache.sqlite3"


def get_cache_connection() -> sqlite3.Connection:
    TRANSLATION_CACHE.parent.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(TRANSLATION_CACHE)
    connection.execute(
        """
        CREATE TABLE IF NOT EXISTS translations (
            language TEXT NOT NULL,
            source TEXT NOT NULL,
            translated TEXT NOT NULL,
            PRIMARY KEY (language, source)
        )
        """
    )
    return connection


def clean_subtitles(subtitles: str) -> str:
    blocks = re.split(r"\r?\n\s*\r?\n", subtitles.strip())
    cleaned_blocks = []

    for block in blocks:
        if PROMOTIONAL_TEXT_PATTERN.search(block):
            continue

        lines = block.splitlines()
        if len(lines) < 3 or "-->" not in lines[1]:
            raise ValueError("O arquivo contém um bloco SRT inválido.")

        caption_text = "\n".join(lines[2:])
        caption_text = MARKUP_TAG_PATTERN.sub("", caption_text)
        caption_text = BRACKETED_DESCRIPTION_PATTERN.sub("", caption_text)

        caption_lines = []
        for line in caption_text.splitlines():
            clean_line = line
            clean_line = clean_line.replace("♪", "").replace("✰", "").strip()
            clean_line = re.sub(r"\s+", " ", clean_line)
            if re.fullmatch(r"[\-\u2013\u2014]\s*", clean_line):
                continue
            if clean_line:
                caption_lines.append(clean_line)

        has_caption = any(
            re.sub(r"^[\-\u2013\u2014]\s*", "", line).strip()
            for line in caption_lines
        )
        if not has_caption:
            continue

        cleaned_blocks.append(
            "\n".join([str(len(cleaned_blocks) + 1), lines[1], *caption_lines])
        )

    return "\n\n".join(cleaned_blocks) + ("\n" if cleaned_blocks else "")


def translate_text(
    text: str,
    target_language: str,
    cancel_callback=None,
) -> str:
    cache_key = text.strip()
    if not cache_key:
        return ""

    with get_cache_connection() as cache:
        cached = cache.execute(
            "SELECT translated FROM translations WHERE language = ? AND source = ?",
            (target_language, cache_key),
        ).fetchone()
    if cached:
        return cached[0]

    params = urlencode(
        {
            "client": "gtx",
            "sl": "auto",
            "tl": target_language,
            "dt": "t",
            "q": text,
        }
    )
    url = f"https://translate.googleapis.com/translate_a/single?{params}"

    last_error: Exception | None = None
    for attempt in range(5):
        if cancel_callback and cancel_callback():
            raise TranslationError("Processo cancelado pelo usuário.")

        try:
            with urlopen(url, timeout=30) as response:
                payload = json.loads(response.read().decode("utf-8"))
            break
        except HTTPError as error:
            last_error = error
            if error.code in (429, 502, 503):
                if error.code == 429:
                    message = "Limite de tradução temporariamente atingido."
                else:
                    message = "O serviço de tradução está temporariamente indisponível."
                if attempt < 4:
                    for _ in range((3 * (attempt + 1)) * 2):
                        if cancel_callback and cancel_callback():
                            raise TranslationError("Processo cancelado pelo usuário.")
                        time.sleep(0.5)
                    continue
                raise TranslationError(
                    f"{message} Aguarde alguns minutos e tente novamente."
                ) from error
            raise TranslationError(
                f"O serviço de tradução respondeu com erro HTTP {error.code}."
            ) from error
        except (URLError, TimeoutError, OSError, UnicodeError, json.JSONDecodeError) as error:
            last_error = error
            if attempt < 4:
                for _ in range((3 * (attempt + 1)) * 2):
                    if cancel_callback and cancel_callback():
                        raise TranslationError("Processo cancelado pelo usuário.")
                    time.sleep(0.5)
                continue
            raise TranslationError(
                "Não foi possível acessar o serviço de tradução. "
                "Verifique sua conexão com a internet."
            ) from error
    else:
        raise TranslationError(
            "Não foi possível concluir a tradução."
        ) from last_error

    translations = payload[0] if payload and isinstance(payload, list) else None
    if not translations:
        raise TranslationError("O serviço de tradução não retornou texto.")

    translated_parts = [
        part[0] for part in translations if isinstance(part, list) and part
    ]
    translated = "".join(translated_parts).strip()
    if not translated:
        raise TranslationError("O serviço de tradução não retornou texto.")

    with get_cache_connection() as cache:
        cache.execute(
            "INSERT OR REPLACE INTO translations(language, source, translated) "
            "VALUES (?, ?, ?)",
            (target_language, cache_key, translated),
        )

    time.sleep(0.2)
    return translated


def translate_subtitles(
    subtitles: str,
    target_language: str,
    provider: str = "google",
    progress_callback=None,
    status_callback=None,
    cancel_callback=None,
) -> str:
    translated_blocks = []
    blocks = subtitles.strip().split("\n\n")
    argos_translation: ArgosTranslation | None = None

    if provider == "argos":
        if target_language != "pt":
            raise TranslationError(
                "A tradução offline do Argos está configurada para português."
            )
        if status_callback:
            status_callback("Carregando o tradutor offline...")
        argos_translation = get_argos_translation(
            status_callback=status_callback,
            cancel_callback=cancel_callback,
        )

    for block_index, block in enumerate(blocks, start=1):
        if cancel_callback and cancel_callback():
            raise TranslationError("Processo cancelado pelo usuário.")
        lines = block.splitlines()
        if len(lines) < 3:
            raise TranslationError("Não foi possível interpretar os blocos da legenda.")

        translated_caption = []
        for line in lines[2:]:
            if not line.strip():
                continue
            if provider == "argos":
                if cancel_callback and cancel_callback():
                    raise TranslationError("Processo cancelado pelo usuário.")
                if argos_translation is None:
                    raise TranslationError("O tradutor offline não foi carregado.")
                translated_caption.append(argos_translation.translate(line))
            else:
                translated_caption.append(
                    translate_text(
                        line,
                        target_language,
                        cancel_callback=cancel_callback,
                    )
                )

        if translated_caption:
            translated_blocks.append(
                "\n".join([lines[0], lines[1], *translated_caption])
            )
        if progress_callback:
            progress_callback(block_index, len(blocks))

    return "\n\n".join(translated_blocks) + (
        "\n" if translated_blocks else ""
    )


def get_argos_translation(
    status_callback=None,
    cancel_callback=None,
) -> ArgosTranslation:
    try:
        argos_package = importlib.import_module("argostranslate.package")
        argos_translate = importlib.import_module("argostranslate.translate")
    except ModuleNotFoundError as error:
        raise TranslationError(
            "O Argos Translate não está instalado. Execute "
            "'python -m pip install -r requirements.txt'."
        ) from error

    def find_translation():
        languages = argos_translate.get_installed_languages()
        source = next((language for language in languages if language.code == "en"), None)
        target = next((language for language in languages if language.code == "pt"), None)
        return source.get_translation(target) if source and target else None

    translation = find_translation()
    if translation:
        return translation

    if cancel_callback and cancel_callback():
        raise TranslationError("Processo cancelado pelo usuário.")
    if status_callback:
        status_callback("Baixando o modelo inglês → português (primeira execução)...")

    try:
        argos_package.update_package_index()
        model = next(
            (
                package
                for package in argos_package.get_available_packages()
                if package.from_code == "en" and package.to_code == "pt"
            ),
            None,
        )
        if model is None:
            raise TranslationError(
                "Não foi encontrado um modelo offline de inglês para português."
            )
        if cancel_callback and cancel_callback():
            raise TranslationError("Processo cancelado pelo usuário.")
        try:
            model_path = model.download()
        except Exception as error:
            raise TranslationError(
                f"Não foi possível baixar o modelo offline: {error}"
            ) from error
        if status_callback:
            status_callback("Instalando o modelo de tradução...")
        argos_package.install_from_path(model_path)
    except TranslationError:
        raise
    except (OSError, RuntimeError, StopIteration) as error:
        raise TranslationError(
            f"Não foi possível instalar o modelo offline: {error}"
        ) from error

    translation = find_translation()
    if translation is None:
        raise TranslationError(
            "O modelo offline não foi instalado corretamente."
        )
    return translation


class SubtitleCleaner:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Corretor de legendas SRT")
        self.root.minsize(800, 500)

        self.output_title = tk.StringVar()
        self.translate_enabled = tk.BooleanVar(value=False)
        self.target_language = tk.StringVar(value="pt-BR")
        self.translation_provider = tk.StringVar(value="Argos (offline)")
        self.source_path: Path | None = None
        self.progress_window: tk.Toplevel | None = None
        self.cancel_processing = threading.Event()
        self._create_widgets()

    def _create_text_area(
        self, parent: ttk.Frame, title: str
    ) -> tk.Text:
        container = ttk.Frame(parent)
        container.grid(row=0, column=0, sticky="nsew")
        parent.rowconfigure(0, weight=1)
        parent.columnconfigure(0, weight=1)

        ttk.Label(container, text=title).grid(
            row=0, column=0, sticky="w", pady=(0, 5)
        )
        text = tk.Text(container, wrap="word", undo=True)
        scrollbar = ttk.Scrollbar(
            container, orient="vertical", command=text.yview
        )
        text.configure(yscrollcommand=scrollbar.set)
        text.grid(row=1, column=0, sticky="nsew")
        scrollbar.grid(row=1, column=1, sticky="ns")
        container.rowconfigure(1, weight=1)
        container.columnconfigure(0, weight=1)
        return text

    def _create_widgets(self) -> None:
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        main = ttk.Frame(self.root, padding=10)
        main.grid(row=0, column=0, sticky="nsew")
        main.columnconfigure((0, 1), weight=1, uniform="text_areas")
        main.rowconfigure(0, weight=1)

        input_frame = ttk.Frame(main)
        input_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 5))
        output_frame = ttk.Frame(main)
        output_frame.grid(row=0, column=1, sticky="nsew", padx=(5, 0))

        self.input_text = self._create_text_area(
            input_frame, "Original: cole ou abra um arquivo .srt"
        )
        self.output_text = self._create_text_area(
            output_frame, "Corrigida: formatação, símbolos e propaganda removidos"
        )
        self.output_text.configure(state="disabled")

        controls = ttk.Frame(main)
        controls.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        controls.columnconfigure(1, weight=1)

        ttk.Button(
            controls, text="Abrir SRT...", command=self.open_file
        ).grid(row=0, column=0, padx=(0, 8))
        self.correct_button = ttk.Button(
            controls, text="Corrigir", command=self.correct
        )
        self.correct_button.grid(row=0, column=1, sticky="w")
        ttk.Checkbutton(
            controls,
            text="Traduzir",
            variable=self.translate_enabled,
        ).grid(row=1, column=0, padx=(0, 8), pady=(8, 0), sticky="w")
        ttk.Label(controls, text="Idioma:").grid(
            row=1, column=1, pady=(8, 0), sticky="w"
        )
        target_combo = ttk.Combobox(
            controls,
            textvariable=self.target_language,
            values=("pt-BR", "en", "es"),
            state="readonly",
            width=8,
        )
        target_combo.grid(
            row=1, column=1, pady=(8, 0), padx=(52, 0), sticky="w"
        )
        ttk.Label(controls, text="Tradutor:").grid(
            row=1, column=2, padx=(10, 5), pady=(8, 0), sticky="w"
        )
        ttk.Combobox(
            controls,
            textvariable=self.translation_provider,
            values=(
                "Argos (offline, inglês → português)",
                "Google (online)",
            ),
            state="readonly",
            width=18,
        ).grid(row=1, column=3, pady=(8, 0), sticky="w")
        ttk.Label(controls, text="Título do arquivo de saída (sem .srt):").grid(
            row=0, column=2, padx=(10, 5)
        )
        ttk.Entry(controls, textvariable=self.output_title).grid(
            row=0, column=3, sticky="ew"
        )
        ttk.Button(
            controls, text="Salvar em Downloads", command=self.save_file
        ).grid(row=0, column=4, padx=(8, 0))

    def open_file(self) -> None:
        selected = filedialog.askopenfilename(
            title="Abrir legenda",
            initialdir=str(Path.home() / "Downloads"),
            filetypes=[("Arquivos SRT", "*.srt"), ("Todos os arquivos", "*.*")],
        )
        if not selected:
            return

        try:
            content = Path(selected).read_text(encoding="utf-8-sig")
        except (OSError, UnicodeError) as error:
            messagebox.showerror("Erro ao abrir arquivo", str(error))
            return

        self.input_text.delete("1.0", "end")
        self.input_text.insert("1.0", content)
        self.output_title.set(Path(selected).stem)
        self.source_path = Path(selected)

    def correct(self) -> None:
        original = self.input_text.get("1.0", "end-1c")
        if not original.strip():
            messagebox.showinfo(
                "Legenda vazia", "Cole ou abra um arquivo SRT antes de corrigir."
            )
            return

        translate_enabled = self.translate_enabled.get()
        target = self.target_language.get()
        provider = (
            "argos"
            if self.translation_provider.get().startswith("Argos")
            else "google"
        )
        self.cancel_processing.clear()
        self.correct_button.configure(state="disabled")
        self._open_progress_window()

        worker = threading.Thread(
            target=self._correct_in_background,
            args=(original, translate_enabled, target, provider),
            daemon=True,
        )
        worker.start()

    def _open_progress_window(self) -> None:
        if self.progress_window is not None:
            return

        window = tk.Toplevel(self.root)
        window.title("Processando legenda")
        window.resizable(False, False)
        window.transient(self.root)
        window.grab_set()

        self.progress_label = ttk.Label(
            window,
            text="Preparando legenda...",
            width=48,
        )
        self.progress_label.grid(row=0, column=0, padx=20, pady=(15, 8))
        self.progress_bar = ttk.Progressbar(
            window,
            orient="horizontal",
            length=360,
            mode="determinate",
        )
        self.progress_bar.grid(row=1, column=0, padx=20, pady=(0, 12))
        ttk.Button(
            window,
            text="Cancelar",
            command=self.cancel_processing.set,
        ).grid(row=2, column=0, pady=(0, 15))
        self.progress_window = window
        window.protocol("WM_DELETE_WINDOW", self.cancel_processing.set)

    def _close_progress_window(self) -> None:
        if self.progress_window is not None:
            self.progress_window.grab_release()
            self.progress_window.destroy()
            self.progress_window = None
        self.correct_button.configure(state="normal")

    def _update_progress(self, current: int, total: int) -> None:
        if self.progress_window is None:
            return
        percentage = int(current * 100 / total) if total else 100
        self.progress_bar.configure(value=percentage)
        self.progress_label.configure(
            text=f"Traduzindo bloco {current} de {total}..."
        )

    def _correct_in_background(
        self,
        original: str,
        translate_enabled: bool,
        target: str,
        provider: str,
    ) -> None:
        try:
            corrected = clean_subtitles(original)
            if translate_enabled:
                if provider == "argos" and target != "pt-BR":
                    raise TranslationError(
                        "O Argos offline está disponível para tradução "
                        "de inglês para português."
                    )
                language_code = "pt" if target == "pt-BR" else target
                corrected = translate_subtitles(
                    corrected,
                    language_code,
                    provider=provider,
                    progress_callback=lambda current, total: self.root.after(
                        0, self._update_progress, current, total
                    ),
                    status_callback=lambda text: self.root.after(
                        0, self._set_progress_message, text
                    ),
                    cancel_callback=self.cancel_processing.is_set,
                )
                if self.cancel_processing.is_set():
                    raise TranslationError("Processo cancelado pelo usuário.")
        except (ValueError, TranslationError) as error:
            self.root.after(0, self._finish_processing, None, str(error))
            return

        self.root.after(0, self._finish_processing, corrected, None)

    def _set_progress_message(self, message: str) -> None:
        if self.progress_window is not None:
            self.progress_label.configure(text=message)

    def _finish_processing(
        self,
        corrected: str | None,
        error_message: str | None,
    ) -> None:
        self._close_progress_window()
        if error_message:
            messagebox.showerror("Erro no processamento", error_message)
            return
        if corrected is None:
            return

        self.output_text.configure(state="normal")
        self.output_text.delete("1.0", "end")
        self.output_text.insert("1.0", corrected)
        self.output_text.configure(state="disabled")

    def save_file(self) -> None:
        content = self.output_text.get("1.0", "end-1c")
        if not content.strip():
            messagebox.showinfo(
                "Nada para salvar", "Corrija a legenda antes de salvar."
            )
            return

        filename = self.output_title.get().strip() or "legendas_corrigidas"
        if not filename.lower().endswith(".srt"):
            filename += ".srt"
        if any(character in filename for character in '<>:"/\\|?*'):
            messagebox.showerror(
                "Nome de arquivo inválido",
                "Informe apenas o nome do arquivo, sem caminho ou caracteres "
                "inválidos.",
            )
            return

        downloads_dir = Path.home() / "Downloads"
        output_path = downloads_dir / filename

        if self.source_path and output_path.resolve() == self.source_path.resolve():
            messagebox.showerror(
                "Nome igual ao arquivo original",
                "Para preservar o SRT original, escolha outro nome para o "
                "arquivo corrigido.",
            )
            return

        if output_path.exists() and not messagebox.askyesno(
            "Substituir arquivo?",
            f"O arquivo já existe:\n{output_path}\n\n"
            "Deseja substituir a versão anterior?",
        ):
            return

        try:
            with open(output_path, "w", encoding="utf-8-sig", newline="\r\n") as output:
                output.write(content)
        except OSError as error:
            messagebox.showerror("Erro ao salvar arquivo", str(error))
            return

        self.output_title.set(output_path.stem)
        messagebox.showinfo(
            "Arquivo salvo", f"Legenda salva em:\n{output_path}"
        )


def main() -> None:
    root = tk.Tk()
    SubtitleCleaner(root)
    root.mainloop()


if __name__ == "__main__":
    main()

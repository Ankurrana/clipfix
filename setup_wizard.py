"""First-run setup wizard for Clipboard Coach."""
import json
import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path

if getattr(sys, "frozen", False):
    APP_DIR = Path(os.environ.get("LOCALAPPDATA", "")) / "ClipboardCoach"
else:
    APP_DIR = Path(__file__).parent
APP_DIR.mkdir(parents=True, exist_ok=True)
CONFIG_FILE = APP_DIR / "config.json"


def run_setup():
    """Show a setup dialog and create config.json. Returns True if configured."""
    if CONFIG_FILE.exists():
        return True

    # Also check env vars
    for key in ("AZURE_OPENAI_API_KEY", "OPENAI_API_KEY", "ANTHROPIC_API_KEY"):
        if os.environ.get(key):
            return True

    root = tk.Tk()
    root.title("Clipboard Coach - Setup")
    root.geometry("520x420")
    root.resizable(False, False)

    # Center on screen
    root.update_idletasks()
    x = (root.winfo_screenwidth() - 500) // 2
    y = (root.winfo_screenheight() - 350) // 2
    root.geometry(f"+{x}+{y}")

    result = {"configured": False}

    # Header
    ttk.Label(root, text="Clipboard Coach Setup", font=("Segoe UI", 14, "bold")).pack(pady=(15, 5))
    ttk.Label(root, text="Choose your LLM provider:", font=("Segoe UI", 10)).pack(pady=(0, 10))

    # Provider selection
    provider_var = tk.StringVar(value="azure_openai")
    providers_frame = ttk.LabelFrame(root, text="Provider", padding=10)
    providers_frame.pack(fill="x", padx=20, pady=5)

    ttk.Radiobutton(providers_frame, text="Azure OpenAI (Azure AI Foundry)", variable=provider_var, value="azure_openai").pack(anchor="w")
    ttk.Radiobutton(providers_frame, text="OpenAI (api.openai.com)", variable=provider_var, value="openai").pack(anchor="w")
    ttk.Radiobutton(providers_frame, text="Anthropic (Claude)", variable=provider_var, value="anthropic").pack(anchor="w")
    ttk.Radiobutton(providers_frame, text="Custom (OpenAI-compatible)", variable=provider_var, value="custom").pack(anchor="w")

    # Config fields
    fields_frame = ttk.LabelFrame(root, text="Configuration", padding=10)
    fields_frame.pack(fill="x", padx=20, pady=5)

    ttk.Label(fields_frame, text="API Key:").grid(row=0, column=0, sticky="w", pady=2)
    api_key_entry = ttk.Entry(fields_frame, width=50, show="*")
    api_key_entry.grid(row=0, column=1, pady=2, padx=(5, 0))

    ttk.Label(fields_frame, text="Model / Deployment:").grid(row=1, column=0, sticky="w", pady=2)
    model_entry = ttk.Entry(fields_frame, width=50)
    model_entry.insert(0, "gpt-4.1")
    model_entry.grid(row=1, column=1, pady=2, padx=(5, 0))

    ttk.Label(fields_frame, text="Endpoint (Azure/Custom):").grid(row=2, column=0, sticky="w", pady=2)
    endpoint_entry = ttk.Entry(fields_frame, width=50)
    endpoint_entry.grid(row=2, column=1, pady=2, padx=(5, 0))

    def save_config():
        provider = provider_var.get()
        api_key = api_key_entry.get().strip()
        model = model_entry.get().strip()
        endpoint = endpoint_entry.get().strip()

        if not api_key:
            messagebox.showerror("Error", "API Key is required.")
            return

        config = {"provider": provider}

        if provider == "azure_openai":
            if not endpoint:
                messagebox.showerror("Error", "Azure endpoint URL is required.")
                return
            config["endpoint"] = endpoint
            config["deployment"] = model or "gpt-4.1"
            config["api_key"] = api_key
            config["api_version"] = "2025-01-01-preview"

        elif provider == "openai":
            config["model"] = model or "gpt-4.1"
            config["api_key"] = api_key

        elif provider == "anthropic":
            config["model"] = model or "claude-sonnet-4-6"
            config["api_key"] = api_key

        elif provider == "custom":
            if not endpoint:
                messagebox.showerror("Error", "Endpoint URL is required for custom provider.")
                return
            config["base_url"] = endpoint
            config["model"] = model
            config["api_key"] = api_key

        CONFIG_FILE.write_text(json.dumps(config, indent=2))
        result["configured"] = True
        root.destroy()

    def cancel():
        root.destroy()

    # Buttons
    btn_frame = ttk.Frame(root)
    btn_frame.pack(pady=10)
    ttk.Button(btn_frame, text="Connect", command=save_config).pack(side="left", padx=5)
    ttk.Button(btn_frame, text="Quit", command=cancel).pack(side="left", padx=5)

    root.mainloop()
    return result["configured"]


if __name__ == "__main__":
    if run_setup():
        print("Configuration saved!")
    else:
        print("Setup cancelled.")

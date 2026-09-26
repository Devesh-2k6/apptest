"use client";

import React, { useState } from "react";
import { Smartphone, QrCode } from "lucide-react";
import { MobileConnectModal } from "./MobileConnectModal";

export function MobileAppFloatingButton() {
  const [isOpen, setIsOpen] = useState(false);

  return (
    <>
      <div className="fixed bottom-20 lg:bottom-6 right-6 z-40">
        <button
          onClick={() => setIsOpen(true)}
          className="group relative flex items-center gap-2 px-3.5 py-2 bg-zinc-900/90 hover:bg-zinc-900 border border-zinc-700/80 hover:border-purple-500/60 text-white rounded-full shadow-lg shadow-purple-950/20 backdrop-blur-xl transition-[transform,border-color,box-shadow] duration-200 cursor-pointer hover:scale-105 active:scale-95 touch-manipulation"
          title="Connect Mobile App (Expo Go / Scanner / APK)"
        >
          {/* Animated ping dot */}
          <span className="relative flex h-2 w-2">
            <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-purple-400 opacity-75" />
            <span className="relative inline-flex rounded-full h-2 w-2 bg-purple-500" />
          </span>

          <Smartphone className="w-3.5 h-3.5 text-purple-400 group-hover:rotate-6 transition-transform duration-200" />
          <span className="text-[11px] font-semibold tracking-tight">
            Mobile App
          </span>
          <QrCode className="w-3 h-3 text-zinc-400 group-hover:text-purple-300 transition-colors" />
        </button>
      </div>

      <MobileConnectModal isOpen={isOpen} onClose={() => setIsOpen(false)} />
    </>
  );
}

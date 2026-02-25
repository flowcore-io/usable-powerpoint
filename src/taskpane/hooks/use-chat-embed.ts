import { RefObject, useEffect, useRef } from "react";
import { UsableChatEmbed } from "../lib/embed-sdk";
import { pptxToolSchemas, handlePptxToolCall } from "../lib/pptx-tools";

// ---------------------------------------------------------------------------
// Configuration
// ---------------------------------------------------------------------------

// Public embed token — configures which Usable workspace/expert is shown.
// Override via Office.context.roamingSettings key "embedTokenOverride".
const DEFAULT_EMBED_TOKEN = "uc_c0d33e3eebfd2e615b6e2d6a4354c22fd10aa8051111f4bfa92b5c0303a342a5";

const IFRAME_ORIGIN = "https://chat.usable.dev";

// ---------------------------------------------------------------------------
// Hook
// ---------------------------------------------------------------------------

/**
 * @param iframeRef        - Ref to the chat iframe element.
 * @param accessToken      - Keycloak JWT to authenticate the embed. Pass null when unauthenticated.
 * @param ensureValidToken - Returns current token if fresh (>60s), otherwise refreshes.
 */
export function useChatEmbed(
  iframeRef: RefObject<HTMLIFrameElement>,
  accessToken: string | null,
  ensureValidToken: () => Promise<string | null>
): void {
  const embedRef = useRef<UsableChatEmbed | null>(null);

  // Ref so the onReady closure always sees the latest token without re-creating the embed.
  const accessTokenRef = useRef<string | null>(accessToken);
  accessTokenRef.current = accessToken;

  // -------------------------------------------------------------------------
  // Create / destroy the embed instance when the iframe mounts
  // -------------------------------------------------------------------------

  useEffect(() => {
    const iframe = iframeRef.current;
    if (!iframe) return;

    // Read embed token override from roaming settings (if set by the user)
    let embedToken = DEFAULT_EMBED_TOKEN;
    try {
      const override = Office.context.roamingSettings.get("embedTokenOverride") as string | null;
      if (override) embedToken = override;
    } catch {
      // roamingSettings may not be available in all contexts
    }

    // Set the iframe src
    const targetSrc = `${IFRAME_ORIGIN}/embed?token=${encodeURIComponent(embedToken)}`;
    if (iframe.src !== targetSrc) {
      iframe.src = targetSrc;
    }

    // Create the embed SDK instance
    const embed = new UsableChatEmbed(iframe, {
      iframeOrigin: IFRAME_ORIGIN,

      onToolCall: async (tool, args, _requestId) => {
        return handlePptxToolCall(tool, args);
      },

      onTokenRefreshRequired: ensureValidToken,

      onError: (code, message) => {
        console.error(`[UsableEmbed] Error ${code}: ${message}`);
      },
    });

    embedRef.current = embed;

    // On READY: register tools AND send the current auth token.
    // Auth must be sent here (not earlier) — postMessages sent before READY are
    // lost because the iframe hasn't set up its listener yet.
    embed.onReady(() => {
      embed.registerTools(pptxToolSchemas);
      if (accessTokenRef.current) {
        embed.setAuth(accessTokenRef.current);
      }
    });

    return () => {
      embed.destroy();
      embedRef.current = null;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [iframeRef]);

  // -------------------------------------------------------------------------
  // Re-send auth whenever the token changes after READY (silent refresh, etc.)
  // -------------------------------------------------------------------------

  useEffect(() => {
    if (accessToken && embedRef.current) {
      // Validate freshness before pushing token to the embed.
      // If the proactive timer fired late (WebView throttling), this
      // will detect near-expiry and fetch a fresh token first.
      ensureValidToken().then((validToken) => {
        if (validToken && embedRef.current) {
          embedRef.current.setAuth(validToken);
        }
      });
    }
  }, [accessToken, ensureValidToken]);
}

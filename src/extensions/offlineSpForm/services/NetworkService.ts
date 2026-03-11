export type NetworkEffectiveType = 'slow-2g' | '2g' | '3g' | '4g';

export interface INetworkState {
  online: boolean;
  isWeak: boolean;
  effectiveType?: NetworkEffectiveType;
  downlinkMbps?: number;
  rttMs?: number;
  saveData?: boolean;
}

interface INetworkInformation {
  effectiveType?: NetworkEffectiveType;
  downlink?: number;
  rtt?: number;
  saveData?: boolean;
  addEventListener?: (type: 'change', listener: () => void) => void;
  removeEventListener?: (type: 'change', listener: () => void) => void;
}

function getNetworkInformation(): INetworkInformation | undefined {
  const nav = navigator as unknown as { connection?: INetworkInformation; mozConnection?: INetworkInformation; webkitConnection?: INetworkInformation };
  return nav.connection || nav.mozConnection || nav.webkitConnection;
}

export class NetworkService {
  public static getCurrentState(): INetworkState {
    const online = navigator.onLine;
    const info = getNetworkInformation();

    const effectiveType = info?.effectiveType;
    const downlinkMbps = typeof info?.downlink === 'number' ? info.downlink : undefined;
    const rttMs = typeof info?.rtt === 'number' ? info.rtt : undefined;
    const saveData = typeof info?.saveData === 'boolean' ? info.saveData : undefined;

    const isWeak = NetworkService._isWeakConnection({ online, effectiveType, downlinkMbps, rttMs, saveData });

    return {
      online,
      isWeak,
      effectiveType,
      downlinkMbps,
      rttMs,
      saveData
    };
  }

  public static subscribe(onChange: (state: INetworkState) => void): () => void {
    const handler = (): void => {
      onChange(NetworkService.getCurrentState());
    };

    window.addEventListener('online', handler);
    window.addEventListener('offline', handler);

    const info = getNetworkInformation();
    info?.addEventListener?.('change', handler);

    // Emit initial state
    handler();

    return () => {
      window.removeEventListener('online', handler);
      window.removeEventListener('offline', handler);
      info?.removeEventListener?.('change', handler);
    };
  }

  private static _isWeakConnection(state: {
    online: boolean;
    effectiveType?: NetworkEffectiveType;
    downlinkMbps?: number;
    rttMs?: number;
    saveData?: boolean;
  }): boolean {
    if (!state.online) return true;

    // If the browser indicates reduced data usage, treat as weak.
    if (state.saveData) return true;

    // Effective type is the clearest signal when available.
    if (state.effectiveType === 'slow-2g' || state.effectiveType === '2g') return true;

    // Heuristics: low throughput or very high latency.
    if (typeof state.downlinkMbps === 'number' && state.downlinkMbps > 0 && state.downlinkMbps < 0.5) return true;
    if (typeof state.rttMs === 'number' && state.rttMs > 1000) return true;

    return false;
  }
}

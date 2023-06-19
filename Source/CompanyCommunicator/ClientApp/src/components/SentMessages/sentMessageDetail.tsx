/* eslint-disable @typescript-eslint/restrict-template-expressions */
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from 'react';
import { useTranslation } from 'react-i18next';
import {
  Badge,
  Button,
  Menu,
  MenuItem,
  MenuList,
  MenuPopover,
  MenuTrigger,
  Table,
  TableBody,
  TableCell,
  TableCellLayout,
  TableHeader,
  TableHeaderCell,
  Caption1,
  TableRow,
  Tooltip,
  useArrowNavigationGroup,
  Body1Strong
} from '@fluentui/react-components';
import {
  CalendarCancel16Regular,
  CalendarCancel24Regular,
  Chat20Regular,
  ChatMultiple24Regular,
  DocumentCopyRegular,
  MoreHorizontal24Filled,
  DismissCircle16Regular,
  Warning16Regular,
  CheckmarkCircle16Regular,
} from '@fluentui/react-icons';
import { dialog, DialogDimension, UrlDialogInfo } from '@microsoft/teams-js';
import { GetDraftMessagesSilentAction, GetSentMessagesSilentAction } from '../../actions';
import { cancelSentNotification, duplicateDraftNotification } from '../../apis/messageListApi';
import { getBaseUrl } from '../../configVariables';
import { formatNumber } from '../../i18n';
import { ROUTE_PARTS, ROUTE_QUERY_PARAMS } from '../../routes';
import { useAppDispatch } from '../../store';

export const SentMessageDetail = (sentMessages: any) => {
  const { t } = useTranslation();
  const keyboardNavAttr = useArrowNavigationGroup({ axis: 'grid' });
  const dispatch = useAppDispatch();
  const statusUrl = (id: string) => getBaseUrl() + `/${ROUTE_PARTS.VIEW_STATUS}/${id}?${ROUTE_QUERY_PARAMS.LOCALE}={locale}`;

  const renderSendingText = (message: any) => {
    let text = '';
    switch (message.status) {
      case 'Queued':
        text = t('Queued');
        break;
      case 'SyncingRecipients':
        text = t('SyncingRecipients');
        break;
      case 'InstallingApp':
        text = t('InstallingApp');
        break;
      case 'Sending':
        // eslint-disable-next-line no-case-declarations, @typescript-eslint/restrict-plus-operands
        const sentCount =
          // eslint-disable-next-line @typescript-eslint/restrict-plus-operands
          (message.succeeded ? message.succeeded : 0) + (message.failed ? message.failed : 0) + (message.unknown ? message.unknown : 0);
        text = t('SendingMessages', {
          SentCount: formatNumber(sentCount),
          TotalCount: formatNumber(message.totalMessageCount),
        });
        break;
      case 'Canceling':
        text = t('Canceling');
        break;
      case 'Canceled':
      case 'Sent':
      case 'Failed':
        text = '';
    }

    return text;
  };

  const shouldNotShowCancel = (msg: any) => {
    let cancelState = false;
    if (msg?.status !== undefined) {
      const status = msg.status.toUpperCase();
      cancelState = status === 'SENT' || status === 'UNKNOWN' || status === 'FAILED' || status === 'CANCELED' || status === 'CANCELING';
    }
    return cancelState;
  };

  const onOpenTaskModule = (event: any, url: string, title: string) => {
    const dialogInfo: UrlDialogInfo = {
      url,
      title,
      size: { height: DialogDimension.Large, width: DialogDimension.Large },
      fallbackUrl: url,
    };

    // now open the dialog
    dialog.url.open(dialogInfo);
  };

  const duplicateDraftMessage = async (id: number) => {
    try {
      await duplicateDraftNotification(id).then(() => {
        GetDraftMessagesSilentAction(dispatch);
      });
    } catch (error) {
      return error;
    }
  };

  const cancelSentMessage = async (id: number) => {
    try {
      await cancelSentNotification(id).then(() => {
        GetSentMessagesSilentAction(dispatch);
      });
    } catch (error) {
      return error;
    }
  };

  return (
    <Table {...keyboardNavAttr} role='grid' className='sent-messages' aria-label={t('sentMessagesGridNavigation') ?? ''}>
      <TableHeader>
        <TableRow>
          <TableHeaderCell key='message' style={{ width: '55%' }}>
            <Body1Strong>{t('message')}</Body1Strong>
          </TableHeaderCell>
          <TableHeaderCell key='recipients'>
            <Body1Strong>{t('Recipients')}</Body1Strong>
          </TableHeaderCell>
          <TableHeaderCell key='actions' style={{ width: '50px' }}>
            <Body1Strong>{t('actions')}</Body1Strong>
          </TableHeaderCell>
        </TableRow>
      </TableHeader>
      <TableBody>
        {sentMessages.sentMessages?.map((item: any) => (
          // eslint-disable-next-line @typescript-eslint/restrict-template-expressions
          <TableRow key={`${item.id}key`}>
            <TableCell tabIndex={0} role='gridcell'>
              <TableCellLayout
                media={<Chat20Regular />}
                title={item.title}
                style={{ cursor: 'pointer' }}
                truncate
                onClick={() => {
                  onOpenTaskModule(null, statusUrl(item.id), t('ViewStatus'));
                }}
              >
                <Body1Strong style={{ whiteSpace: 'nowrap' }}>{item.title}</Body1Strong>
                {renderSendingText(item) && <><br /><Badge size='small' appearance="tint" color="warning">{renderSendingText(item)}</Badge></>}
                {item.sentDate && <><br /><Badge size='small' appearance="tint" color="informative">{item.sentDate}</Badge></>}
                <br />
                <Caption1>{item.createdBy}</Caption1>
              </TableCellLayout>
            </TableCell>
            <TableCell tabIndex={0} role='gridcell'>
              <TableCellLayout>
                <Tooltip content={t('TooltipSuccess') ?? ''} relationship='label'>
                  <span style={{ paddingLeft: '2px' }}>
                    <Badge appearance="tint" icon={<CheckmarkCircle16Regular />} color="success">{formatNumber(item.succeeded)}</Badge>
                  </span>
                </Tooltip>
                <Tooltip content={t('TooltipFailure') ?? ''} relationship='label'>
                  <span style={{ paddingLeft: '2px' }}>
                    <Badge appearance="tint" icon={<DismissCircle16Regular />} color="severe">{formatNumber(item.failed)}</Badge>
                  </span>
                </Tooltip>
                {item.canceled && (
                  <>
                    <Tooltip content='Canceled' relationship='label'>
                      <span style={{ paddingLeft: '2px' }}>
                        <Badge appearance="tint" icon={<CalendarCancel16Regular />} color="danger">{formatNumber(item.canceled)}</Badge>
                      </span>
                    </Tooltip>
                  </>
                )}
                {item.unknown && (
                  <>
                    <Tooltip content='Unknown' relationship='label'>
                      <span style={{ paddingLeft: '2px' }}>
                        <Badge appearance="tint" icon={<Warning16Regular />} color="warning">{formatNumber(item.unknown)}</Badge>
                      </span>
                    </Tooltip>
                  </>
                )}
              </TableCellLayout>
            </TableCell>
            <TableCell role='gridcell' style={{ width: '50px' }}>
              <TableCellLayout style={{ float: 'right' }}>
                <Menu>
                  <MenuTrigger disableButtonEnhancement>
                    <Button aria-label='Actions menu' icon={<MoreHorizontal24Filled />} />
                  </MenuTrigger>
                  <MenuPopover>
                    <MenuList>
                      <MenuItem
                        icon={<ChatMultiple24Regular />}
                        key={'viewStatusKey'}
                        onClick={() => {
                          onOpenTaskModule(null, statusUrl(item.id), t('ViewStatus'));
                        }}
                      >
                        {t('ViewStatus')}
                      </MenuItem>
                      {
                        // eslint-disable-next-line @typescript-eslint/no-misused-promises, @typescript-eslint/promise-function-async
                        <MenuItem key={'duplicateKey'} icon={<DocumentCopyRegular />} onClick={() => duplicateDraftMessage(item.id)}>
                          {t('Duplicate')}
                        </MenuItem>
                      }
                      {!shouldNotShowCancel(item) && (
                        // eslint-disable-next-line @typescript-eslint/no-misused-promises, @typescript-eslint/promise-function-async
                        <MenuItem key={'cancelKey'} icon={<CalendarCancel24Regular />} onClick={() => cancelSentMessage(item.id)}>
                          {t('Cancel')}
                        </MenuItem>
                      )}
                    </MenuList>
                  </MenuPopover>
                </Menu>
              </TableCellLayout>
            </TableCell>
          </TableRow>
        ))}
      </TableBody>
    </Table>
  );
};
